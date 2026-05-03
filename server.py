"""Local Flask bridge for opening Outlook draft emails from the web form."""

from __future__ import annotations

import logging
import os
import re
import socket
import tempfile
import threading
import uuid
import webbrowser
from pathlib import Path
from typing import Iterable

from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

try:
    import pythoncom
    import win32com.client as win32_client
except ImportError:
    pythoncom = None
    win32_client = None


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = Path(tempfile.mkdtemp(prefix="outlook_drafts_"))
ALLOWED_STATIC_FILES = {"index.html", "style.css", "app.js"}
ALLOWED_ATTACHMENT_EXTENSIONS = {".pdf", ".doc", ".docx"}
EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)
log.info("Temporary attachments folder: %s", UPLOAD_FOLDER)

app = Flask(__name__)
CORS(app)


def port_is_busy(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.35)
        return sock.connect_ex((host, port)) == 0


def choose_port(preferred_port: int) -> int:
    if not port_is_busy("127.0.0.1", preferred_port):
        return preferred_port

    for port in range(preferred_port + 1, preferred_port + 20):
        if not port_is_busy("127.0.0.1", port):
            return port

    raise RuntimeError("לא נמצא פורט פנוי להפעלת השרת המקומי.")


def normalize_recipients(values: Iterable[str]) -> list[str]:
    """Accept repeated form fields and comma/semicolon/newline separated values."""
    recipients: list[str] = []
    seen: set[str] = set()

    for value in values:
        for part in re.split(r"[,;\n\r\t ]+", value):
            email = part.strip()
            key = email.lower()
            if email and key not in seen:
                recipients.append(email)
                seen.add(key)

    return recipients


def save_attachment() -> tuple[Path | None, str | None]:
    uploaded_file = request.files.get("cv_file")
    if not uploaded_file or not uploaded_file.filename:
        return None, None

    original_path = Path(uploaded_file.filename)
    extension = original_path.suffix.lower()
    if extension not in ALLOWED_ATTACHMENT_EXTENSIONS:
        return None, "ניתן לצרף קובץ PDF, DOC או DOCX בלבד"

    safe_stem = secure_filename(original_path.stem) or "cv_attachment"
    original_name = f"{safe_stem}{extension}"
    unique_name = f"{uuid.uuid4().hex}_{original_name}"
    attachment_path = UPLOAD_FOLDER / unique_name
    uploaded_file.save(attachment_path)
    log.info("Saved attachment: %s", attachment_path)
    return attachment_path, None


def outlook_unavailable_message() -> str | None:
    if os.name != "nt":
        return "הפתרון דורש Windows, כי Outlook נפתח דרך COM Automation."
    if pythoncom is None or win32_client is None:
        return "חסרה החבילה pywin32. הריצי: python -m pip install -r requirements.txt"
    return None


@app.get("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.get("/<path:filename>")
def static_files(filename: str):
    if filename not in ALLOWED_STATIC_FILES:
        return jsonify({"success": False, "error": "הקובץ לא נמצא"}), 404
    return send_from_directory(BASE_DIR, filename)


@app.get("/health")
def health():
    unavailable = outlook_unavailable_message()
    return jsonify(
        {
            "app": "outlook-drafts-local",
            "success": unavailable is None,
            "status": "ready" if unavailable is None else "limited",
            "message": unavailable or "השרת המקומי פעיל.",
        }
    )


@app.post("/create-drafts")
def create_drafts():
    subject = request.form.get("subject", "").strip()
    body = request.form.get("body", "").strip()
    recipients = normalize_recipients(request.form.getlist("recipients"))

    if not subject:
        return jsonify({"success": False, "error": "חסר נושא מייל"}), 400
    if not recipients:
        return jsonify({"success": False, "error": "חסרה לפחות כתובת נמען אחת"}), 400
    if not body:
        return jsonify({"success": False, "error": "חסר גוף הודעה"}), 400

    invalid = [email for email in recipients if not EMAIL_RE.match(email)]
    if invalid:
        return jsonify({"success": False, "error": f"כתובת מייל לא תקינה: {invalid[0]}"}), 400

    unavailable = outlook_unavailable_message()
    if unavailable:
        return jsonify({"success": False, "error": unavailable}), 500

    attachment_path, attachment_error = save_attachment()
    if attachment_error:
        return jsonify({"success": False, "error": attachment_error}), 400

    created = 0
    errors: list[dict[str, str]] = []

    try:
        pythoncom.CoInitialize()
        outlook = win32_client.Dispatch("Outlook.Application")

        for recipient in recipients:
            try:
                mail = outlook.CreateItem(0)  # 0 = olMailItem
                mail.To = recipient
                mail.Subject = subject
                mail.Body = body

                if attachment_path and attachment_path.exists():
                    mail.Attachments.Add(str(attachment_path))

                mail.Display(False)
                created += 1
                log.info("Draft opened for %s", recipient)
            except Exception as exc:  # Outlook COM errors should not hide other recipients.
                log.exception("Failed to create draft for %s", recipient)
                errors.append({"recipient": recipient, "error": str(exc)})
    except Exception as exc:
        log.exception("Could not start Outlook")
        return jsonify({"success": False, "error": f"לא ניתן לפתוח את Outlook: {exc}"}), 500
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    if created == 0:
        return jsonify({"success": False, "error": "לא נפתחה אף טיוטה", "details": errors}), 500

    return jsonify({"success": True, "created": created, "errors": errors})


if __name__ == "__main__":
    preferred = int(os.environ.get("PORT", "5000"))
    port = choose_port(preferred)
    url = f"http://127.0.0.1:{port}/"

    print("=" * 58)
    print("שרת מקומי לפתיחת טיוטות Outlook פעיל")
    print(f"פתחי בדפדפן: {url}")
    print("לעצירה: Ctrl+C")
    print("=" * 58)

    if os.environ.get("OPEN_BROWSER", "1") != "0":
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    app.run(host="127.0.0.1", port=port, debug=False)
