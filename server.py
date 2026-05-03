"""Local Flask bridge for opening Outlook draft emails from the web form."""

from __future__ import annotations

import json
import logging
import os
import re
import socket
import subprocess
import sys
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
    import pythoncom  # noqa: F401
    import win32com.client  # noqa: F401
except ImportError:
    pythoncom = None
    win32com = None


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = Path(tempfile.mkdtemp(prefix="outlook_drafts_"))
BRIDGE_TIMEOUT_SECONDS = 45
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

    raise RuntimeError("No free local port was found for the Flask server.")


def normalize_recipients(values: Iterable[str]) -> list[str]:
    """Accept repeated fields and comma/semicolon/newline separated values."""
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
        return None, "Only PDF, DOC, and DOCX attachments are supported."

    safe_stem = secure_filename(original_path.stem) or "cv_attachment"
    unique_name = f"{uuid.uuid4().hex}_{safe_stem}{extension}"
    attachment_path = UPLOAD_FOLDER / unique_name
    uploaded_file.save(attachment_path)
    log.info("Saved attachment: %s", attachment_path)
    return attachment_path, None


def outlook_unavailable_message() -> str | None:
    if os.name != "nt":
        return "This solution requires Windows because Outlook is controlled through COM Automation."
    if pythoncom is None:
        return "The pywin32 package is missing. Run: python -m pip install -r requirements.txt"
    return None


def run_outlook_bridge(
    subject: str,
    body: str,
    recipients: list[str],
    attachment_path: Path | None,
) -> tuple[dict, int]:
    payload = {
        "subject": subject,
        "body": body,
        "recipients": recipients,
        "attachment_path": str(attachment_path) if attachment_path else None,
    }
    payload_path = UPLOAD_FOLDER / f"{uuid.uuid4().hex}_payload.json"
    payload_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    try:
        completed = subprocess.run(
            [sys.executable, str(BASE_DIR / "outlook_bridge.py"), str(payload_path)],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=BRIDGE_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return (
            {
                "success": False,
                "created": 0,
                "errors": [],
                "error": (
                    "Outlook did not respond in time. Open classic Outlook first "
                    "and close any profile, login, or security dialog."
                ),
            },
            504,
        )
    finally:
        try:
            payload_path.unlink()
        except OSError:
            pass

    stdout = completed.stdout.strip()
    stderr = completed.stderr.strip()
    log.info("Outlook bridge exit code: %s", completed.returncode)
    if stderr:
        log.error("Outlook bridge stderr: %s", stderr)

    try:
        result = json.loads(stdout) if stdout else {}
    except json.JSONDecodeError:
        result = {
            "success": False,
            "created": 0,
            "errors": [],
            "error": "Outlook bridge returned an invalid response.",
            "stdout": stdout,
            "stderr": stderr,
        }

    if completed.returncode != 0 and not result.get("error"):
        result["error"] = stderr or "Outlook bridge failed."

    return result, 200 if result.get("success") else 500


@app.get("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.get("/<path:filename>")
def static_files(filename: str):
    if filename not in ALLOWED_STATIC_FILES:
        return jsonify({"success": False, "error": "File not found."}), 404
    return send_from_directory(BASE_DIR, filename)


@app.get("/health")
def health():
    unavailable = outlook_unavailable_message()
    return jsonify(
        {
            "app": "outlook-drafts-local",
            "success": unavailable is None,
            "status": "ready" if unavailable is None else "limited",
            "message": unavailable or "Local server is ready.",
        }
    )


@app.post("/create-drafts")
def create_drafts():
    subject = request.form.get("subject", "").strip()
    body = request.form.get("body", "").strip()
    recipients = normalize_recipients(request.form.getlist("recipients"))

    if not subject:
        return jsonify({"success": False, "error": "Email subject is required."}), 400
    if not recipients:
        return jsonify({"success": False, "error": "At least one recipient is required."}), 400
    if not body:
        return jsonify({"success": False, "error": "Email body is required."}), 400

    invalid = [email for email in recipients if not EMAIL_RE.match(email)]
    if invalid:
        return jsonify({"success": False, "error": f"Invalid email address: {invalid[0]}"}), 400

    unavailable = outlook_unavailable_message()
    if unavailable:
        return jsonify({"success": False, "error": unavailable}), 500

    attachment_path, attachment_error = save_attachment()
    if attachment_error:
        return jsonify({"success": False, "error": attachment_error}), 400

    result, status_code = run_outlook_bridge(subject, body, recipients, attachment_path)
    return jsonify(result), status_code


if __name__ == "__main__":
    preferred = int(os.environ.get("PORT", "5000"))
    port = choose_port(preferred)
    url = f"http://127.0.0.1:{port}/"

    print("=" * 58)
    print("Outlook draft local server is running")
    print(f"Open in browser: {url}")
    print("Stop with Ctrl+C")
    print("=" * 58)

    if os.environ.get("OPEN_BROWSER", "1") != "0":
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
