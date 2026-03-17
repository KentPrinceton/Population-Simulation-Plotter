## Front-end asset structure

This dashboard now keeps presentation and behavior code outside the main HTML file.

### Folders

- `css/dashboard.css`
  - Theme tokens, layout styles, and component styling.
- `js/dashboard/app.js`
  - Dashboard application logic (state, parsing, plotting, BE analysis, exports).

### Why this structure

- Reduces merge conflicts in the main HTML file.
- Makes CSS and JavaScript updates independent.
- Keeps future refactors straightforward (e.g., splitting `app.js` into feature modules).

