let API_BASE = "";
let serverReady = false;

type StatusType = "success" | "error" | "info";

function getEl<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id) as T | null;
  if (!el) {
    throw new Error(`Element with id="${id}" was not found`);
  }
  return el;
}

const form = getEl<HTMLFormElement>("mailForm");
const fileInput = getEl<HTMLInputElement>("cvFile");
const dropZone = getEl<HTMLDivElement>("dropZone");
const fileNameEl = getEl<HTMLDivElement>("fileName");
const serverStatus = getEl<HTMLDivElement>("serverStatus");
const submitBtn = getEl<HTMLButtonElement>("submitBtn");

getEl<HTMLButtonElement>("addRecipient").addEventListener("click", () => {
  const wrapper = getEl<HTMLDivElement>("recipientsWrapper");
  const row = document.createElement("div");
  row.className = "recipient-row";
  row.innerHTML = `
    <input type="email" placeholder="email@example.com" class="recipient-input" autocomplete="email" />
    <button type="button" class="btn-icon remove-recipient" title="הסר נמען" aria-label="הסר נמען">&times;</button>
  `;
  wrapper.appendChild(row);
  updateRemoveButtons();
  row.querySelector<HTMLInputElement>(".recipient-input")?.focus();
});

getEl<HTMLDivElement>("recipientsWrapper").addEventListener("click", (event: MouseEvent) => {
  const target = event.target as HTMLElement;
  if (target.classList.contains("remove-recipient")) {
    target.closest(".recipient-row")?.remove();
    updateRemoveButtons();
  }
});

function updateRemoveButtons(): void {
  const rows = document.querySelectorAll<HTMLDivElement>(".recipient-row");
  rows.forEach((row) => {
    const button = row.querySelector<HTMLButtonElement>(".remove-recipient");
    if (button) {
      button.hidden = rows.length === 1;
    }
  });
}

fileInput.addEventListener("change", () => {
  showFileName(fileInput.files?.[0] ?? null);
});

dropZone.addEventListener("dragover", (event: DragEvent) => {
  event.preventDefault();
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (event: DragEvent) => {
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

form.addEventListener("submit", async (event: SubmitEvent) => {
  event.preventDefault();

  if (!serverReady) {
    await checkServer();
  }

  if (!serverReady) {
    showStatus("השרת המקומי לא פעיל. הפעילי את start_server.bat ונסי שוב.", "error");
    return;
  }

  const recipients = getRecipients();
  const subject = getEl<HTMLInputElement>("subject").value.trim();
  const body = getEl<HTMLTextAreaElement>("body").value.trim();
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
    const data = (await response.json()) as { success: boolean; created?: number; error?: string };

    if (!response.ok || !data.success) {
      showStatus(`שגיאה: ${data.error ?? "לא ניתן לפתוח טיוטות"}`, "error");
      return;
    }

    showStatus(`נפתחו ${data.created ?? recipients.length} טיוטות ב-Outlook.`, "success");
    resetForm();
  } catch {
    showStatus("לא ניתן להתחבר לשרת המקומי. הפעילי את start_server.bat ונסי שוב.", "error");
  } finally {
    submitBtn.disabled = false;
  }
});

function getRecipients(): string[] {
  const recipients: string[] = [];
  const inputs = document.querySelectorAll<HTMLInputElement>(".recipient-input");

  inputs.forEach((input: HTMLInputElement) => {
    input.value.split(/[,\s;]+/).forEach((value: string) => {
      const email = value.trim();
      if (email) {
        recipients.push(email);
      }
    });
  });

  return recipients;
}

function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function showFileName(file: File | null): void {
  fileNameEl.textContent = file ? `נבחר קובץ: ${file.name}` : "";
}

function showStatus(message: string, type: StatusType): void {
  const status = getEl<HTMLDivElement>("statusMsg");
  status.textContent = message;
  status.className = `status-msg ${type}`;
}

function resetForm(): void {
  form.reset();
  fileNameEl.textContent = "";
  getEl<HTMLDivElement>("recipientsWrapper").innerHTML = `
    <div class="recipient-row">
      <input type="email" placeholder="email@example.com" class="recipient-input" autocomplete="email" required />
      <button type="button" class="btn-icon remove-recipient" title="הסר נמען" aria-label="הסר נמען">&times;</button>
    </div>
  `;
  updateRemoveButtons();
}

async function checkServer(): Promise<boolean> {
  serverReady = false;

  for (const baseUrl of getServerCandidates()) {
    try {
      const response = await fetch(`${baseUrl}/health`);
      const data = (await response.json()) as { app?: string; success: boolean };

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

function getServerCandidates(): string[] {
  const candidates: string[] = [];
  const seen = new Set<string>();

  function addCandidate(url: string): void {
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
