# Event QR Check-in System

Single-file web app (`index.html`) + Google Apps Script backend (`Code.gs`).
All devices share one live list and one check-in record through your Google Sheet.

## 1. Backend (Google Sheet)
1. Open the attendee sheet. Row 1 must be headers: a token/QR column, name column(s), an email column.
2. Extensions > Apps Script. Paste `Code.gs` (keep your QR-generation code; read the note at the top about `doGet`).
3. Deploy > Manage deployments > pencil > New version > Deploy.
   Execute as **Me**, access **Anyone**. The URL stays the same.
4. The script adds three columns automatically: `Checked In`, `Check-in Time`, `Checked In By`.

## 2. Frontend (GitHub Pages)
1. Upload `index.html` to the repo root.
2. Settings > Pages > Deploy from a branch > main > / (root).

## 3. Use
- Open the HTTPS link on each device. Scanning the first QR link (or pasting the Web App URL in Setup) connects it to the sheet.
- Every device pulls the attendee list and all check-ins every 30 seconds (tap the status badge to sync now).
- A second scan of the same person on any device shows "Already checked in".
- Offline scans are saved on the device and pushed to the sheet when the connection returns.
