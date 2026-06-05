# ExpenseFlow Family Finance

A GitHub Pages-ready family finance tracker based on your Excel workbook idea.

## Features

- Monthly dashboard with month selector
- Daily expense add/edit/delete
- Month-wise filtering and reports
- Partner tracking: Person A and Person P
- Income tracking for both partners
- Budget vs actual by category
- Add/rename/delete budget items
- Loan tracker
- Prepayment options with one-tap add as expense
- Surplus allocation: 40% liquidity and 60% investment boost
- Month-over-month comparison
- CSV export
- JSON backup and restore
- Print / Save PDF monthly report
- Offline support after first load

## How to host on GitHub Pages

1. Create a GitHub repository, for example `expenseflow`.
2. Upload these files to the repository root:
   - `index.html`
   - `style.css`
   - `app.js`
   - `manifest.json`
   - `sw.js`
3. Go to repository **Settings**.
4. Open **Pages**.
5. Under **Build and deployment**, select:
   - Source: `Deploy from a branch`
   - Branch: `main`
   - Folder: `/root`
6. Save.
7. Open the GitHub Pages URL on Android Chrome.
8. Tap Chrome menu ⋮ → **Add to Home screen**.

## Important data note

This app stores data in your browser using `localStorage`.
Use **Backup JSON** regularly.
If you clear browser data, the app data can be deleted unless you have a backup.

## Best use flow

1. Set Partner names and incomes.
2. Check/edit Budget items.
3. Add loan details and prepayment options.
4. Add expenses daily.
5. Use Dashboard and Reports monthly.
