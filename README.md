# SBO Bid Tracker — Free iPhone web app (Version 1)

This is a **static**, browser-based prototype. It does not contact the Kalitta website, send your files to a server, or store employee credentials. It uses JavaScript Excel and PDF parsing libraries loaded from public CDNs. The company files are read locally in your browser's memory.

## Use it

1. Create a free GitHub account, create a **public repository containing only these app files** (no company data), and enable GitHub Pages from the main branch root. The source code is public; do **not** commit any bid spreadsheets, roster data, credentials, or personal information.
2. Open the resulting GitHub Pages link on your iPhone. In Safari choose Share → Add to Home Screen.
3. On the Files tab, select the 777 monthly lines PDF, monthly Bid List workbook (777 worksheet), and the Excel export from Reports → View Other Bids → View All.
4. Tap Analyse. The projection appears on Home and searchable pilots appear on Pilots. Adjust your preference order on My bids.

## Important

- This is a prototype. Actual report structures can differ; if import validation fails, the parser must be adjusted using a **redacted** sample. The Excel export is expected to have `Name`, `Choice`, `Line` column headings.
- Exclusions are drawn **only** from the `FO CQ SCHEDULED` and `VAC FO` lists on the 777 worksheet of the monthly Bid List workbook, never from separate vacation documents.
- The denominator of `Fxx - a/b` is interpreted as the number of FO slots for the line.
- Awards are simulated from the most senior down, each receiving the first ranked line with a remaining FO position. Blank bids remain unassigned. This does not guarantee the airline's official award.
- The files are **not saved or synchronized**; re-import after a page refresh or on another device. Never upload company data to the GitHub repository.
- Github Pages is public hosting. It hosts the app **code only**, not any crew data. The site is not password protected. Ask your employer before processing or displaying other pilots' information on personal devices.
- Requires an internet connection to load Excel and PDF libraries from CDN. It has no automated company website downloads.
