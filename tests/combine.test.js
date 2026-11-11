/* =====================================================================
   Automated tests for the combining logic.
   Run:  node --test tests/
   They build the sample files, push them through the same functions the
   page uses, write a real .xlsx, then read it back and check it.
   ===================================================================== */
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const XLSX = require("../vendor/xlsx.full.min.js");
const Core = require("../js/combine-core.js");
const { writeSamples, OUT_DIR } = require("./make-samples.js");

const CONFIG = { REQUIRE_LOGIN: false, MAX_FILES: 50, MAX_FILE_MB: 25, OUTPUT_PREFIX: "Combined" };
const ILLEGAL = /[\[\]:*?\/\\]/;

// Write the samples to disk, then read them back like a browser would.
const manifest = writeSamples();
const sampleFiles = manifest.map((m) => ({
  name: m.name,                                             // the intended name (may contain ":")
  bytes: new Uint8Array(fs.readFileSync(path.join(OUT_DIR, m.file)))
}));

// The same steps js/app.js takes when files are dropped in.
function loadFiles(files, config = CONFIG) {
  const entries = [];
  const skipped = [];
  for (const file of files) {
    const verdict = Core.checkIncomingFile({ name: file.name, size: file.bytes.length }, entries, config);
    if (!verdict.ok) { skipped.push(verdict.reason); continue; }
    let entry;
    try {
      entry = Core.parseFile(XLSX, file.bytes, file.name);
    } catch (err) {
      skipped.push(err.message);
      continue;
    }
    if (entry.rowCount === 0) { skipped.push('Skipped "' + file.name + '": there is no data in it.'); continue; }
    entries.push(entry);
    Core.assignTabNames(entries);
  }
  return { entries, skipped };
}

// Build the workbook, turn it into .xlsx bytes, and read those bytes back in.
function roundTrip(entries, options) {
  const built = Core.buildWorkbook(XLSX, entries, options);
  const bytes = Core.workbookToBytes(XLSX, built.workbook);
  const outFile = path.join(OUT_DIR, "..", "combined-" + (options && options.addIndex ? "index" : "plain") +
    (options && options.keepAsText === false ? "-numbers" : "-text") + ".xlsx");
  fs.writeFileSync(outFile, bytes);
  return { built, bytes, workbook: XLSX.read(bytes, { type: "array" }) };
}

function sheetRows(workbook, tabName) {
  const sheet = workbook.Sheets[tabName];
  if (!sheet["!ref"]) return [];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "", blankrows: true });
}

const loaded = loadFiles(sampleFiles);
const VALID_COUNT = sampleFiles.length - 1;                 // everything except empty.csv

test("1. leading zeros survive with 'keep as text' ON", () => {
  const { workbook } = roundTrip(loaded.entries, { keepAsText: true });
  const a = workbook.Sheets["points_a"];
  assert.deepEqual({ t: a.A2.t, v: a.A2.v }, { t: "s", v: "001" });
  assert.deepEqual({ t: a.A4.t, v: a.A4.v }, { t: "s", v: "0042" });
  assert.deepEqual({ t: a.A6.t, v: a.A6.v }, { t: "s", v: "00123" });
  // long coordinates keep every digit, and stay text
  assert.deepEqual({ t: a.B6.t, v: a.B6.v }, { t: "s", v: "6255100.1234567" });
  assert.equal(a.D4.v, "513.000");                          // trailing zeros kept too
  // "1/2" must not become a date
  const b = workbook.Sheets["points_b"];
  assert.deepEqual({ t: b.A5.t, v: b.A5.v }, { t: "s", v: "1/2" });
  assert.equal(b.A2.v, "0501");
  // tab-separated file
  assert.equal(workbook.Sheets["tab_separated"].A2.v, "007");
  assert.equal(workbook.Sheets["tab_separated"].E2.v, "BM, BRASS CAP");
});

test("1b. quoted commas, doubled quotes, accents and the BOM are handled", () => {
  const { workbook } = roundTrip(loaded.entries, { keepAsText: true });
  const q = workbook.Sheets["quoted_commas_bom"];
  assert.equal(q.A1.v, "Point");                            // no invisible BOM stuck to the first cell
  assert.equal(q.A2.v, "0001");
  assert.equal(q.B2.v, "WALL, NORTH FACE");
  assert.equal(q.B3.v, 'ANCHOR 3/4" DIA, EPOXY');
  assert.equal(q.C2.v, "José");
  assert.equal(q.C3.v, "Zoë");
});

test("1c. with 'keep as text' OFF, numbers become real numbers", () => {
  const { workbook } = roundTrip(loaded.entries, { keepAsText: false });
  const a = workbook.Sheets["points_a"];
  assert.equal(a.A2.t, "n");
  assert.equal(a.A2.v, 1);
  assert.equal(a.E2.v, "CP NAIL");
  assert.equal(workbook.Sheets["real"].B2.v, 100.5);
});

test("2. every tab name is valid: 31 characters or fewer, no illegal characters, unique", () => {
  const { workbook } = roundTrip(loaded.entries, { keepAsText: true, addIndex: true });
  const seen = new Set();
  for (const name of workbook.SheetNames) {
    assert.ok(name.length >= 1 && name.length <= 31, "bad length: " + name);
    assert.ok(!ILLEGAL.test(name), "illegal character in: " + name);
    assert.ok(!name.startsWith("'") && !name.endsWith("'"), "apostrophe at the edge of: " + name);
    assert.ok(!seen.has(name.toLowerCase()), "duplicate tab: " + name);
    seen.add(name.toLowerCase());
  }
  assert.deepEqual(Core.validateTabNames(workbook.SheetNames).filter(Boolean), []);
  // the specific tricky ones
  const names = loaded.entries.map((e) => e.tabName);
  assert.ok(names.includes("big_names_and_bad_characters -t"), names.join(" | "));
  assert.ok(names.includes("duplicate name"));
  assert.ok(names.includes("duplicate name (2)"));
});

test("2b. tab-name cleaning rules, one by one", () => {
  assert.equal(Core.cleanTabName("points.csv"), "points");
  assert.equal(Core.cleanTabName("a[b]c:d*e?f/g\\h.xlsx"), "a-b-c-d-e-f-g-h");
  assert.equal(Core.cleanTabName("'quoted'.csv"), "quoted");
  assert.equal(Core.cleanTabName(".csv"), "Sheet");
  assert.equal(Core.cleanTabName("   .csv"), "Sheet");
  assert.equal(Core.cleanTabName("History.csv"), "History-1");
  assert.equal(Core.cleanTabName("x".repeat(80) + ".csv").length, 31);

  // uniqueness is case-insensitive and still fits in 31 characters
  const used = {};
  assert.equal(Core.makeUniqueName("Points", used), "Points");
  assert.equal(Core.makeUniqueName("points", used), "points (2)");
  assert.equal(Core.makeUniqueName("POINTS", used), "POINTS (3)");
  const long = "y".repeat(31);
  assert.equal(Core.makeUniqueName(long, used), long);
  const second = Core.makeUniqueName(long, used);
  assert.equal(second.length, 31);
  assert.ok(second.endsWith(" (2)"));

  // what the live check says while someone is typing a name
  assert.equal(Core.tabNameProblem("Fine name"), "");
  assert.notEqual(Core.tabNameProblem(""), "");
  assert.notEqual(Core.tabNameProblem("   "), "");
  assert.notEqual(Core.tabNameProblem("x".repeat(32)), "");
  assert.notEqual(Core.tabNameProblem("a/b"), "");
  assert.notEqual(Core.tabNameProblem("'start"), "");
  assert.notEqual(Core.tabNameProblem("end'"), "");
  const dupes = Core.validateTabNames(["Alpha", "Beta", "ALPHA"]);
  assert.equal(dupes[0], "");
  assert.equal(dupes[1], "");
  assert.notEqual(dupes[2], "");

  // the workbook builder refuses bad names instead of writing a broken file
  const copy = loaded.entries.slice(0, 2).map((e) => Object.assign({}, e, { tabName: "same" }));
  assert.throws(() => Core.buildWorkbook(XLSX, copy, {}), /same name/);
});

test("3. tab count equals the number of valid files, in input order", () => {
  assert.equal(loaded.entries.length, VALID_COUNT);
  const { workbook } = roundTrip(loaded.entries, { keepAsText: true });
  assert.equal(workbook.SheetNames.length, VALID_COUNT);
  assert.deepEqual(workbook.SheetNames, loaded.entries.map((e) => e.tabName));
  assert.deepEqual(loaded.entries.map((e) => e.fileName),
    sampleFiles.filter((f) => f.name !== "empty.csv").map((f) => f.name));

  // reordering the list reorders the tabs
  const reversed = loaded.entries.slice().reverse();
  const again = roundTrip(reversed, { keepAsText: true }).workbook;
  assert.deepEqual(again.SheetNames, reversed.map((e) => e.tabName));
});

test("3b. the Index tab goes first and lists every file", () => {
  const { workbook } = roundTrip(loaded.entries, { keepAsText: true, addIndex: true });
  assert.equal(workbook.SheetNames.length, VALID_COUNT + 1);
  assert.equal(workbook.SheetNames[0], "Index");
  const rows = sheetRows(workbook, "Index");
  assert.equal(rows.length, VALID_COUNT + 1);
  assert.deepEqual(rows[0], ["#", "File name", "Tab name", "Rows", "Columns"]);
  assert.deepEqual(rows[1], [1, "points_a.csv", "points_a", 6, 5]);

  // if a file is already called Index, the Index tab steps aside
  const clash = loadFiles([{ name: "Index.csv", bytes: new TextEncoder().encode("a,b\n1,2\n") }]);
  const built = Core.buildWorkbook(XLSX, clash.entries, { addIndex: true });
  assert.deepEqual(built.tabNames, ["Index (2)", "Index"]);
});

test("4. the empty file is skipped with a warning", () => {
  assert.ok(!loaded.entries.some((e) => e.fileName === "empty.csv"));
  assert.equal(loaded.skipped.length, 1);
  assert.match(loaded.skipped[0], /empty\.csv/);
  assert.match(loaded.skipped[0], /empty/);
});

test("4b. other files that should be skipped are skipped, with a reason", () => {
  const enc = (s) => new TextEncoder().encode(s);
  const good = { name: "good.csv", bytes: enc("a,b\n1,2\n") };

  // unsupported type
  let r = loadFiles([{ name: "photo.png", bytes: enc("not a sheet") }, good]);
  assert.equal(r.entries.length, 1);
  assert.match(r.skipped[0], /isn't supported/);

  // duplicate (same name and size)
  r = loadFiles([good, good]);
  assert.equal(r.entries.length, 1);
  assert.match(r.skipped[0], /already in the list/);

  // too big (limit shrunk so the test stays fast)
  r = loadFiles([good], Object.assign({}, CONFIG, { MAX_FILE_MB: 0.000001 }));
  assert.equal(r.entries.length, 0);
  assert.match(r.skipped[0], /bigger than/);

  // too many
  r = loadFiles([good, { name: "b.csv", bytes: enc("a\n1\n") }, { name: "c.csv", bytes: enc("a\n1\n") }],
    Object.assign({}, CONFIG, { MAX_FILES: 2 }));
  assert.equal(r.entries.length, 2);
  assert.match(r.skipped[0], /limit is 2 files/);

  // a fake .xlsx gets a plain-English error, not a crash
  r = loadFiles([{ name: "fake.xlsx", bytes: enc("this is not a real workbook") }]);
  assert.equal(r.entries.length, 0);
  assert.match(r.skipped[0], /looks corrupted or isn't really a spreadsheet: fake\.xlsx/);

  // a file with only blank space has no data
  r = loadFiles([{ name: "blank.csv", bytes: enc("\n") }]);
  assert.equal(r.entries.length, 0);
});

test("5. a multi-sheet xlsx only contributes its first sheet", () => {
  const real = loaded.entries.find((e) => e.fileName === "real.xlsx");
  assert.equal(real.sheetCount, 2);
  assert.equal(real.sheetName, "FirstSheet");
  assert.deepEqual(real.warnings, ["Only the first sheet (FirstSheet) was used."]);

  const { workbook } = roundTrip(loaded.entries, { keepAsText: true });
  assert.equal(workbook.SheetNames.filter((n) => n.toLowerCase().includes("sheet")).length, 0);
  const rows = sheetRows(workbook, "real");
  assert.deepEqual(rows[0], ["Point", "Easting", "Northing"]);
  assert.deepEqual(rows[1], ["A1", "100.5", "200.25"]);
  assert.ok(!JSON.stringify(rows).includes("ignored"));
});

test("6. row and column counts in the output match the inputs", () => {
  const expected = {
    "points_a.csv": [6, 5],
    "points_b.csv": [5, 5],
    "tab_separated.tsv": [3, 5],
    "big_names_and_bad_characters [test]:1.csv": [3, 2],
    "duplicate name.csv": [2, 2],
    "duplicate name.tsv": [3, 2],
    "real.xlsx": [4, 3],
    "quoted_commas_bom.csv": [4, 3]
  };
  for (const keepAsText of [true, false]) {
    const { workbook } = roundTrip(loaded.entries, { keepAsText });
    for (const entry of loaded.entries) {
      const [rows, cols] = expected[entry.fileName];
      assert.equal(entry.rowCount, rows, entry.fileName + " rows (input)");
      assert.equal(entry.colCount, cols, entry.fileName + " columns (input)");
      const range = XLSX.utils.decode_range(workbook.Sheets[entry.tabName]["!ref"]);
      assert.equal(range.e.r + 1, rows, entry.fileName + " rows (output, text=" + keepAsText + ")");
      assert.equal(range.e.c + 1, cols, entry.fileName + " columns (output, text=" + keepAsText + ")");
    }
  }
});

test("7. extras: frozen top row, sensible column widths, and the file name", () => {
  const { bytes, built } = roundTrip(loaded.entries, { keepAsText: true });

  // every sheet inside the .xlsx has the "freeze row 1" instruction
  const zip = XLSX.CFB.read(bytes, { type: "array" });
  let sheetsChecked = 0;
  zip.FullPaths.forEach((p, i) => {
    if (!/xl\/worksheets\/sheet\d+\.xml$/.test(p)) return;
    const xml = new TextDecoder().decode(new Uint8Array(zip.FileIndex[i].content));
    assert.match(xml, /<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"\/>/);
    sheetsChecked++;
  });
  assert.equal(sheetsChecked, VALID_COUNT);

  // widths: never wider than 60, never silly-narrow
  const wide = Core.parseFile(XLSX, new TextEncoder().encode("a,b\n" + "z".repeat(200) + ",1\n"), "wide.csv");
  Core.assignTabNames([wide]);
  const cols = Core.buildWorkbook(XLSX, [wide], {}).workbook.Sheets["wide"]["!cols"];
  assert.equal(cols[0].wch, 60);
  assert.ok(cols[1].wch >= 6);
  for (const name of built.tabNames) {
    for (const col of built.workbook.Sheets[name]["!cols"]) assert.ok(col.wch >= 6 && col.wch <= 60);
  }

  // Combined_YYYY-MM-DD.xlsx using the local date
  assert.equal(Core.outputFileName("Combined", new Date(2026, 10, 11, 23, 59)), "Combined_2026-11-11.xlsx");
  assert.equal(Core.outputFileName("Combined", new Date(2027, 0, 5, 0, 1)), "Combined_2027-01-05.xlsx");
});
