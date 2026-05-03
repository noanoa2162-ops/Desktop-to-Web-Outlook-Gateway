"use strict";

let API_BASE = "";
let serverReady = false;

function getEl(id) {
  const el = document.getElementById(id);
  if (!el) {
    throw new Error(`Element with id="${id}" was not found`);
  }
  return el;
}

const form = getEl("mailForm");
const fileInput = getEl("cvFile");
const dropZone = getEl("dropZone");
const fileNameEl = getEl("fileName");
const serverStatus = getEl("serverStatus");
const submitBtn = getEl("submitBtn");

getEl("addRecipient").addEventListener("click", () => {
  const wrapper = getEl("recipientsWrapper");
  const row = document.createElement("div");
  row.className = "recipient-row";
  row.innerHTML = `
    <input type="email" placeholder="email@example.com" class="recipient-input" autocomplete="email" />
    <button type="button" class="btn-icon remove-recipient" title="הסר נמען" aria-label="הסר נמען">&times;</button>
  `;
  wrapper.appendChild(row);
  updateRemoveButtons();
  row.querySelector(".recipient-input").focus();
});

getEl("recipientsWrapper").addEventListener("click", (event) => {
  const target = event.target;
  if (target.classList.contains("remove-recipient")) {
    target.closest(".recipient-row").remove();
    updateRemoveButtons();
  }
});

function updateRemoveButtons() {
  const rows = document.querySelectorAll(".recipient-row");
  rows.forEach((row) => {
    const button = row.querySelector(".remove-recipient");
    if (button) {
      button.hidden = rows.length === 1;
    }
  });
}

fileInput.addEventListener("change", () => {
  showFileName(fileInput.files?.[0] ?? null);
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragover");

  const file = event.dataTransfer?.files?.[0] ?? null;
  if (!file) {
    return;
  }

  const dataTransfer = new DataTransfer();
  dataTransfer.items.add(file);
  fileInput.files = dataTransfer.files;
  showFileName(file);
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!serverReady) {
    await checkServer();
  }

  if (!serverReady) {
    showStatus("השרת המקומי לא פעיל. הפעילי את start_server.bat ונסי שוב.", "error");
    return;
  }

  const recipients = getRecipients();
  const subject = getEl("subject").value.trim();
  const body = getEl("body").value.trim();
  const file = fileInput.files?.[0];

  if (!subject) {
    showStatus("נא למלא נושא מייל.", "error");
    return;
  }

  if (recipients.length === 0) {
    showStatus("נא להזין לפחות כתובת נמען אחת.", "error");
    return;
  }

  const invalidRecipient = recipients.find((email) => !isValidEmail(email));
  if (invalidRecipient) {
    showStatus(`כתובת מייל לא תקינה: ${invalidRecipient}`, "error");
    return;
  }

  if (!body) {
    showStatus("נא למלא גוף הודעה.", "error");
    return;
  }

  submitBtn.disabled = true;
  showStatus("פותח טיוטות ב-Outlook...", "info");

  try {
    const formData = new FormData();
    formData.append("subject", subject);
    formData.append("body", body);
    recipients.forEach((recipient) => formData.append("recipients", recipient));
    if (file) {
      formData.append("cv_file", file);
    }

    const response = await fetch(`${API_BASE}/create-drafts`, {
      method: "POST",
      body: formData,
    });
    const data = await response.json();

    if (!response.ok || !data.success) {
      showStatus(`שגיאה: ${data.error ?? "לא ניתן לפתוח טיוטות"}`, "error");
      return;
    }

    showStatus(`נפתחו ${data.created} טיוטות ב-Outlook.`, "success");
    resetForm();
  } catch {
    showStatus("לא ניתן להתחבר לשרת המקומי. הפעילי את start_server.bat ונסי שוב.", "error");
  } finally {
    submitBtn.disabled = false;
  }
});

function getRecipients() {
  const recipients = [];
  const inputs = document.querySelectorAll(".recipient-input");

  inputs.forEach((input) => {
    input.value.split(/[,\s;]+/).forEach((value) => {
      const email = value.trim();
      if (email) {
        recipients.push(email);
      }
    });
  });

  return recipients;
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function showFileName(file) {
  fileNameEl.textContent = file ? `נבחר קובץ: ${file.name}` : "";
}

function showStatus(message, type) {
  const status = getEl("statusMsg");
  status.textContent = message;
  status.className = `status-msg ${type}`;
}

function resetForm() {
  form.reset();
  fileNameEl.textContent = "";
  getEl("recipientsWrapper").innerHTML = `
    <div class="recipient-row">
      <input type="email" placeholder="email@example.com" class="recipient-input" autocomplete="email" required />
      <button type="button" class="btn-icon remove-recipient" title="הסר נמען" aria-label="הסר נמען">&times;</button>
    </div>
  `;
  updateRemoveButtons();
}

async function checkServer() {
  serverReady = false;

  for (const baseUrl of getServerCandidates()) {
    try {
      const response = await fetch(`${baseUrl}/health`);
      const data = await response.json();

      if (data.app !== "outlook-drafts-local") {
        continue;
      }

      API_BASE = baseUrl;
      serverReady = data.success === true;
      serverStatus.textContent = data.success ? "שרת פעיל" : "שרת מוגבל";
      serverStatus.className = data.success ? "server-status ready" : "server-status limited";
      return serverReady;
    } catch {
      // Try the next local port.
    }
  }

  serverStatus.textContent = "שרת לא פעיל";
  serverStatus.className = "server-status offline";
  return false;
}

function getServerCandidates() {
  const candidates = [];
  const seen = new Set();

  function addCandidate(url) {
    if (!seen.has(url)) {
      candidates.push(url);
      seen.add(url);
    }
  }

  if (window.location.protocol.startsWith("http")) {
    addCandidate("");
  }

  for (let port = 5000; port <= 5019; port += 1) {
    addCandidate(`http://127.0.0.1:${port}`);
  }

  return candidates;
}

updateRemoveButtons();
checkServer();
