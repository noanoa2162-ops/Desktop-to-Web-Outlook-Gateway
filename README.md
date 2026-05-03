# Desktop-to-Web Outlook Gateway

Desktop-to-Web Outlook Gateway is a local web application for recruitment workflows. It allows a recruiter to enter candidate email content in a browser and generate ready-to-review Outlook draft emails on the user's desktop.

The system is designed for a common staffing-agency scenario: a recruiter needs to send a candidate profile and CV to several companies, but the final email should still be reviewed, edited, and sent manually from the recruiter's personal Outlook account.

## What The Application Does

The application provides a browser-based form for preparing candidate submission emails:

- Email subject
- One or more recipient addresses
- Message body
- Candidate CV attachment, as a PDF or Word document

When the form is submitted, the application creates a separate Outlook draft for each recipient. Each draft includes the entered subject, message body, recipient address, and attached CV file. The system does not send emails automatically; it opens editable drafts so the recruiter stays in control of the final review and send action.

## User Experience

The interface is intentionally focused and operational. It presents a single workflow: prepare an email, attach the CV, and open Outlook drafts.

The form supports multiple recipients while keeping each outgoing message separate. This avoids creating one shared group email and gives the recruiter a clean draft per company or hiring contact.

## Architecture

The project combines a browser UI with a small local backend service:

```text
Browser UI
  index.html
  style.css
  app.js

Local Python Bridge
  server.py
  Flask
  outlook_bridge.ps1
  PowerShell / Outlook COM Automation

Desktop Application
  Microsoft Outlook
```

The browser cannot directly control local desktop software such as Outlook. For security reasons, websites are not allowed to freely access local files, launch desktop programs, or use Windows COM APIs.

To solve that limitation, the project uses a local Flask server as a controlled bridge. The browser sends the form data to `localhost`, and the Python service delegates Outlook automation to a short-lived PowerShell bridge process. That process uses Outlook COM Automation to create draft email windows in the installed desktop Outlook application.

## Request Flow

```text
Recruiter fills the web form
        |
        v
Frontend sends FormData to the local Flask server
        |
        v
Flask validates recipients, subject, body, and attachment
        |
        v
The uploaded CV is saved temporarily on the local machine
        |
        v
The Outlook bridge process runs PowerShell COM Automation
        |
        v
One editable draft email is created per recipient
```

## Why Not Use `mailto:`

The project intentionally avoids relying on `mailto:` links because they are not sufficient for this workflow.

`mailto:` can open a basic email compose window, but it does not reliably support file attachments, has compatibility differences between clients and browsers, and is not suitable for automatically creating multiple separate drafts with a CV attached.

The local bridge approach gives the application more control while still keeping the final send action manual and user-approved.

## Key Technical Decisions

- **Flask backend:** provides a lightweight local API between the browser and the desktop environment.
- **PowerShell COM bridge:** enables Outlook desktop automation on Windows.
- **Separate drafts:** each recipient receives an individual draft instead of being grouped into one email.
- **Manual sending:** the app prepares drafts only; it does not send messages automatically.
- **Local-only operation:** the bridge runs on `127.0.0.1`, keeping the desktop automation scoped to the user's machine.

## Project Structure

```text
index.html         Main web form
style.css          UI styling
app.js             Browser-side form logic and API calls
app.ts             TypeScript source version of the browser logic
server.py          Local Flask server and Outlook automation bridge
outlook_bridge.ps1 Short-lived Outlook COM automation process
requirements.txt   Python dependencies
start_server.bat   Windows launcher for the local service
tsconfig.json      TypeScript configuration
```

## Platform Scope

This solution targets Windows machines with classic Microsoft Outlook installed and configured. The Outlook automation layer depends on COM APIs, which are available in the classic desktop version of Outlook.

The newer "New Outlook" client does not expose the same COM automation interface, so the desktop bridge approach is intended for environments where classic Outlook is available.

## Security And Control

The application is built around a deliberate safety boundary:

- The browser UI collects the email content.
- The local bridge handles desktop automation.
- Outlook opens drafts for the user to inspect.
- The user manually decides whether to edit, send, or discard each message.

This design avoids silent email sending and keeps the recruiter in control of communication from their personal mailbox.
