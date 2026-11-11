# CSV Combiner: session log

## 2026-11-11: first build (Stages 0 to 5 of the pipeline)

Built from `Mr-Code-CSV-Combiner-Pipeline.md`.

### What was built
A standalone, browser-only web app. You drop in CSV / TSV / TXT / XLS / XLSX / XLSM / ODS files and
download one Excel workbook with one tab per file. Nothing is uploaded anywhere.

### Preflight (Stage 0)
- git 2.55, node 24.18, npm 11.16: all present.
- GitHub CLI (`gh`): present and logged in as `robbiefromramona-press`.
- Netlify CLI: **not installed**. Plan: link the GitHub repo in the Netlify dashboard (the usual pattern for these tools).
- `press-csv-combiner` folder did not exist yet. Good.
- `bim-press` homepage is `WEBSITE/index.html`. The CSV Converter and Field Calculator cards live in the
  "Apps and Tools" section, in the extra row that the **View All Tools** button reveals (`id="moreTools"`).
  Colors and fonts were copied from there (orange `#ff6b1a` on near-black, Barlow / Barlow Semi Condensed).

### Files created
`index.html`, `css/styles.css`, `js/config.js`, `js/combine-core.js`, `js/app.js`,
`vendor/xlsx.full.min.js` (SheetJS 0.20.3), `tests/make-samples.js`, `tests/combine.test.js`,
`tests/index.js`, `netlify.toml`, `README.md`, `SESSION_LOG.md`, `.gitignore`.

### Decisions made along the way (and why)
- **Logic split into `js/combine-core.js`** so the same code runs on the page and in the tests.
- **Text files are decoded by us, then handed to SheetJS as text.** Reading a CSV as raw bytes mangles
  accented letters when there is no BOM marker. We try UTF-8 first and fall back to Windows-1252.
- **"Keep every value as text" writes every cell as text** for all file types, not just CSVs.
- **Frozen top row is added by hand.** The free edition of SheetJS cannot write freeze panes, so after it
  builds the .xlsx we open it (it is a zip) and add the one line per sheet that freezes row 1. Because of
  that the download uses a normal browser download instead of `XLSX.writeFile`. If the patch ever fails,
  the file still downloads, just without the frozen row.
- **Up / down arrows next to the drag handle.** Drag and drop does not work on phones or by keyboard; the arrows do.
- **A file called .xlsx / .xlsm / .ods that is not really one** gets the plain-English "looks corrupted" message
  (SheetJS would otherwise happily read the junk as text).
- **Files with only blank lines** count as empty and are skipped.
- **Windows cannot have `:` in a file name**, so the sample `big_names_and_bad_characters [test]:1.csv` is saved
  with `_` on disk; `tests/output/samples/manifest.json` keeps the intended name and the tests and the page check use it.
- **`tests/index.js`** exists because Node 22+ reads `node --test tests/` as "run tests/index.js".
- "History" is also blocked as a tab name (Excel reserves it).

### Test results
`node --test tests/`: **12 passed, 0 failed.** Covers: leading zeros kept, quoted commas / BOM / accents,
numbers mode, tab-name rules, tab count and order, Index tab, empty and other skipped files, first-sheet-only,
row and column counts, frozen row, column widths, output file name.

Local browser check (Stage 4): loaded all samples plus a .png and a duplicate, renamed a tab (bad name blocks
the button with a red message), reordered by arrows and by drag, opened a preview, combined with the Index tab on,
read the downloaded workbook back: 9 tabs in the right order, `001` still `001`. No console errors. Checked at phone width.

### Not done yet
- Stage 6: GitHub repo + Netlify site (waiting for Robbie's yes).
- Stage 7: card on bim-press.com.
- The result has not been opened in real Excel by a human yet (verified by reading the file back with SheetJS and a zip integrity check).

### Phase 2 ideas
- Send straight to Google Sheets.
- Login gate via Netlify Identity (the `REQUIRE_LOGIN` switch and TODO hook are already in place).
- "All sheets" option for multi-sheet workbooks.
- "Stack into one sheet" mode that matches columns across files.
