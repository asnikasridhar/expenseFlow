IMPORTANT FIX

This package fixes the Excel data import bug.

What was wrong:
- The previous app.js referred to EXCEL_IMPORTED_EXPENSES before it was declared.
- That can stop JavaScript from running, so the dashboard stays ₹0.

Upload/replace these files in GitHub root:
- index.html
- style.css
- app.js
- manifest.json
- sw.js
- README.md

After upload:
1. Open https://asnikasridhar.github.io/expenseFlow/?v=2
2. Hard refresh.
3. On Android Chrome: Settings > Site settings > All sites > github.io > Clear data if old values still show.

Expected imported data:
- May 2026: 136 entries, ₹4,22,038
- June 2026: 21 entries, ₹1,43,443
- Total: 157 entries
