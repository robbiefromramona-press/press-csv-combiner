# CSV Combiner (BIM-Press)

Drop in a pile of CSV or spreadsheet files and download **one Excel workbook (.xlsx) with one tab per file**.

Everything happens in the visitor's browser. Files are never uploaded to any server.

- Live site: https://csv-combine.netlify.app (the Netlify project is named `csv-combine`)
- Part of the BIM-Press tool family: https://bim-press.com and https://totalstationtech.com

## What it does

- Accepts `.csv`, `.tsv`, `.txt`, `.xls`, `.xlsx`, `.xlsm`, `.ods` (up to 50 files, 25 MB each).
- Each file becomes one tab. Only the **first sheet** of a multi-sheet workbook is used (with a warning).
- Tab names are cleaned up automatically so Excel accepts them (31 characters max, no `[ ] : * ? / \`, no duplicates) and can be edited.
- Drag the handle or use the arrows to reorder the tabs.
- **Keep every value as text** (on by default) protects point numbers like `001` and long coordinates from being changed by Excel.
- Optional **Index** tab listing every file, tab name, and row and column counts.
- Top row is frozen and columns are auto-sized on every tab.

## Files

| File | What it is |
| --- | --- |
| `index.html` | The page. |
| `css/styles.css` | All the styling (BIM-Press orange and black). |
| `js/config.js` | Settings: file limits, output name, and the login on/off switch. |
| `js/combine-core.js` | The "brains": tab-name cleaning, reading files, building the workbook. No page code, so it can be tested. |
| `js/app.js` | The page code: drop zone, file list, buttons. |
| `vendor/xlsx.full.min.js` | SheetJS, the library that reads and writes spreadsheets. |
| `tests/` | Automated tests and the sample-file maker. |
| `netlify.toml` | Tells Netlify to serve this folder as-is (no build step) with basic security headers. |

## SheetJS version

`vendor/xlsx.full.min.js` is **SheetJS Community Edition 0.20.3**, downloaded 2026-11-11 from
`https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js`.

It is saved inside this repo on purpose, so the site keeps working even if that CDN has a bad day.
Do not swap it for the `xlsx` package on npm: that one is old and no longer updated.

## Turning on a login later

In `js/config.js`, change `REQUIRE_LOGIN: false` to `true`. The tool is then replaced by a
"Please log in via BIM-Press" panel. The real login check is not built yet: look for
`// TODO: Netlify Identity check goes here` near the top of `js/app.js`.

## Running the tests

```
node --test tests/
```

This creates sample files in `tests/output/samples/` (handy for trying the app by hand too),
combines them, writes real `.xlsx` files to `tests/output/`, reads them back and checks them.
No `npm install` is needed.

To only make the sample files: `node tests/make-samples.js`

## Running it locally

It is a plain static site, so any static server works, for example:

```
python -m http.server 8809
```

then open http://localhost:8809

## Deploying

Netlify is linked to this GitHub repo. Every push to `main` goes live automatically.
Publish directory is the repo root and there is no build command.
