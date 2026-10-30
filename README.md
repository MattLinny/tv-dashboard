# Serve at Oceanside: volunteer sign-ups

One page, two views. Edit `CONFIG` at the top of the script in `index.html`.

- **Phone page** `https://mattlinny.github.io/tv-dashboard/`: every team and position with open spots. Tap **Join** → "Sign me up" → **I'm in**. The spot fills straight away.
- **Lobby TV** `https://mattlinny.github.io/tv-dashboard/tv.html` (same as `/?tv`): portrait, full screen, a QR code to the phone page, every spot as a filled name or "Open". Checks the sheet every 15 s; new names pop.

## Sheet
`apps-script/Code.gs` turns a Google Sheet into the backend. Setup steps are at the top of that file. Then put the web app URL in `CONFIG.apiUrl`.

- **Teams** tab: one row per position: `team, position, needed, lead_email`. Rows show in sheet order. `needed` = how many people. `lead_email` (optional) also gets an email for each sign-up.
- **Signups** tab: `submitted_at, team, position, name, email, phone, notes, status`. Set `status` to `removed` (or delete the row) to free a spot.
- The page only receives team, position, needed and first name + last initial. Emails, phones and notes stay in the sheet.
- `seedDemo` fills about half the spots with sample names for testing; `clearSignups` empties the Signups tab before launch.

Without `apiUrl` the page shows sample teams and sign-ups only stay in that browser.

## TV setup
1. Laptop → TV over HDMI, laptop on power, auto-sleep off.
2. Rotate the display in the OS (Mac: System Settings → Displays → Rotation 90°). If you can't, use `tv.html?rotate`.
3. Chrome full screen (Cmd+Ctrl+F). Scan the QR once from a phone before the service.

## Look
Oceanside Blue field, Cream cards, Ink Blue text, Sky for open spots, Rust once on "Scan to join".
Type: Perfectly Nineties when it's installed on the computer showing the page, otherwise Instrument Serif (free, from Google Fonts) + Rethink Sans.
Textures (`tex/`) are baked from Matt's scans: Blue Cardboard (book cloth) for the blue, Fine Paper (1) for the cream.
