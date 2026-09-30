# Mystical Beasts Alliance — Google Sheets Connected

This version is configured with the user's published Google Sheets CSV URL.

Google Sheets CSV:
https://docs.google.com/spreadsheets/d/e/2PACX-1vS6ATu1MX6uUYNJg7gZjS8FRMqFLNwc2GZyHsejsJbuzEENQzwD8Tf-ObUy8mp0drVau6S2xLt090Jr/pub?gid=0&single=true&output=csv

How it works:
1. Open the web app.
2. The app attempts to fetch the latest BEASTS data from Google Sheets.
3. Press “↻ ซิงก์ข้อมูลจาก Google Sheets” to refresh manually.
4. If Google Sheets cannot be reached, the app falls back to BEASTS.csv bundled with the site.

For GitHub Pages:
- Upload the files in this folder to the repository root.
- Enable Settings > Pages > Deploy from branch > main > /(root).
- Open the generated Pages URL.

Note: The remote Google Sheets connection can only be verified after the site is served over HTTP/HTTPS. Opening index.html directly from a file:// URL may be blocked by browser CORS rules.
