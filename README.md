# Event QR Check-in System

Single-file web app (`index.html`) for event registration and QR check-in.

## Deploy on GitHub Pages
1. Create a new repository on GitHub (public).
2. Upload `index.html` (and this README) to the repository root.
3. Go to **Settings > Pages**.
4. Under **Build and deployment**, set Source to **Deploy from a branch**, choose **main** and **/ (root)**, then Save.
5. After about a minute your site is live at `https://<username>.github.io/<repo-name>/`.

GitHub Pages serves over HTTPS, which is required for live camera scanning.

## Using the scanner
- Open the HTTPS link in Chrome or Safari (not inside WhatsApp/Instagram in-app browsers).
- Tap **Start Camera** and allow camera access.
- If live scanning is blocked, use **Scan from Photo** as a fallback.
