# Apps Script Snippets

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-V8-4285F4?logo=google&logoColor=white)
![License: MIT](https://img.shields.io/badge/license-MIT-green)

Small, reusable Google Apps Script utilities for Google Sheets. Each one came out of a real school-operations need and has been generalized so you can drop it into any spreadsheet.

| Snippet | What it does |
|---|---|
| [`split-rows-into-tabs`](snippets/split-rows-into-tabs) | Reads a list from a **Source** sheet and creates one tab per category, for example a training catalog grouped by scope. It includes a menu and a sample-data generator. |
| [`restructure-sheet-columns`](snippets/restructure-sheet-columns) | A safe, **re-runnable migration** for multi-tab workbooks. It drops columns by header name, renames a header on every tab, and creates a new tab from a template with an extra column and dropdowns. |

## Using a snippet

1. Open your Google Sheet → **Extensions → Apps Script**.
2. Paste the snippet's `Code.gs`.
3. Edit the config block at the top of the file.
4. Run the main function from the editor, or use the custom menu where one is provided.

Every snippet is idempotent: running it twice won't duplicate tabs or delete the wrong columns.

## License

[MIT](LICENSE) © Randa ELGawish

## Author

**Randa ELGawish**, Academic leader and EdTech developer
[LinkedIn](https://www.linkedin.com/in/randaelgawishegy) · [GitHub](https://github.com/RandaELGawish93)
