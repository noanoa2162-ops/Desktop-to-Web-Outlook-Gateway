"""Create Outlook draft messages from a JSON payload."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pythoncom
import win32com.client as win32_client


def debug(message: str) -> None:
    print(f"[outlook_bridge] {message}", file=sys.stderr, flush=True)


def create_drafts(payload: dict) -> dict:
    subject = payload["subject"]
    body = payload["body"]
    recipients = payload["recipients"]
    attachment_path = payload.get("attachment_path")

    if attachment_path and not Path(attachment_path).exists():
        return {
            "success": False,
            "created": 0,
            "error": "Attachment file was not found on the local machine.",
            "errors": [],
        }

    created = 0
    errors: list[dict[str, str]] = []

    debug("CoInitialize")
    pythoncom.CoInitialize()
    try:
        debug("Dispatch Outlook.Application")
        outlook = win32_client.Dispatch("Outlook.Application")
        debug("Outlook dispatch ready")

        for recipient in recipients:
            try:
                debug(f"Create item for {recipient}")
                mail = outlook.CreateItem(0)  # 0 = olMailItem
                debug("Set fields")
                mail.To = recipient
                mail.Subject = subject
                mail.Body = body

                if attachment_path:
                    debug(f"Attach file {attachment_path}")
                    mail.Attachments.Add(os.path.abspath(attachment_path))

                debug("Display draft")
                mail.Display(False)
                debug("Draft displayed")
                created += 1
            except Exception as exc:
                debug(f"Error for {recipient}: {exc}")
                errors.append({"recipient": recipient, "error": str(exc)})
    finally:
        debug("CoUninitialize")
        pythoncom.CoUninitialize()

    return {
        "success": created > 0,
        "created": created,
        "errors": errors,
        "error": "" if created > 0 else "No Outlook draft was created.",
    }


def main() -> int:
    if len(sys.argv) != 2:
        print(json.dumps({"success": False, "error": "Missing payload path."}))
        return 2

    payload_path = Path(sys.argv[1])
    payload = json.loads(payload_path.read_text(encoding="utf-8-sig"))
    result = create_drafts(payload)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("success") else 1


if __name__ == "__main__":
    raise SystemExit(main())
