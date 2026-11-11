/* =====================================================================
   CSV Combiner: the "brains".
   ---------------------------------------------------------------------
   This file has NO page code in it (no buttons, no drag and drop).
   It only knows how to:
     1. clean up tab names so Excel accepts them
     2. read one spreadsheet file into rows
     3. glue many files into one workbook
   Because it never touches the page, the same file runs in the browser
   AND in the automated tests (node --test tests/).
   It needs SheetJS (vendor/xlsx.full.min.js), passed in as "XLSX".
   ===================================================================== */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();          // Node (the tests)
  } else {
    root.CombinerCore = factory();       // Browser (the page)
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var MAX_TAB_LEN = 31;                                  // Excel's hard limit
  var ILLEGAL_CHARS = /[\[\]:*?\/\\]/g;                  // [ ] : * ? / \
  var ALLOWED_EXTENSIONS = ["csv", "tsv", "txt", "xls", "xlsx", "xlsm", "ods"];
  var TEXT_EXTENSIONS = ["csv", "tsv", "txt"];
  var MAX_COL_WIDTH = 60;
  var MIN_COL_WIDTH = 6;

  /* ---------- small helpers ---------- */

  function getExtension(fileName) {
    var m = /\.([A-Za-z0-9]{1,5})$/.exec(String(fileName || ""));
    return m ? m[1].toLowerCase() : "";
  }

  function stripExtension(fileName) {
    return String(fileName || "").replace(/\.[A-Za-z0-9]{1,5}$/, "");
  }

  function toBytes(data) {
    if (data instanceof Uint8Array) return data;
    return new Uint8Array(data);                         // ArrayBuffer -> bytes
  }

  /* ---------- tab names ---------- */

  // Turn any file name into a tab name Excel will accept.
  function cleanTabName(rawName) {
    var name = stripExtension(rawName);
    name = name.replace(ILLEGAL_CHARS, "-");
    name = name.replace(/[\u0000-\u001f]/g, " ");        // stray control characters
    name = name.replace(/\s+/g, " ").trim();
    name = name.replace(/^'+|'+$/g, "").trim();          // no apostrophe at either end
    if (name.length > MAX_TAB_LEN) {
      name = name.slice(0, MAX_TAB_LEN).trim().replace(/'+$/g, "").trim();
    }
    if (!name) name = "Sheet";
    if (name.toLowerCase() === "history") name = "History-1";   // Excel reserves "History"
    return name;
  }

  // If "name" is already taken, add " (2)", " (3)"... and still fit in 31 characters.
  // usedLower is a plain object used as a set of lower-case names; it gets updated.
  function makeUniqueName(name, usedLower) {
    var candidate = name;
    var n = 2;
    while (usedLower[candidate.toLowerCase()]) {
      var suffix = " (" + n + ")";
      candidate = name.slice(0, MAX_TAB_LEN - suffix.length).trim() + suffix;
      n++;
    }
    usedLower[candidate.toLowerCase()] = true;
    return candidate;
  }

  // Check one name on its own. Returns "" if fine, or a plain-English problem.
  function tabNameProblem(name) {
    name = String(name == null ? "" : name);
    if (!name.trim()) return "Tab name can't be blank.";
    if (name.length > MAX_TAB_LEN) return "Tab name is too long (" + name.length + " characters, Excel allows 31).";
    if (/[\[\]:*?\/\\]/.test(name)) return "Tab name can't contain any of these: [ ] : * ? / \\";
    if (name.charAt(0) === "'" || name.charAt(name.length - 1) === "'") return "Tab name can't start or end with an apostrophe.";
    if (name.toLowerCase() === "history") return "Excel reserves the name \"History\". Pick another.";
    return "";
  }

  // Check a whole list of names. Returns a list the same length:
  // "" where the name is fine, otherwise the problem (including duplicates).
  function validateTabNames(names) {
    var seen = {};
    return names.map(function (name) {
      var problem = tabNameProblem(name);
      if (problem) return problem;
      var key = String(name).toLowerCase();
      if (seen[key]) return "Two tabs can't share the same name (\"" + name + "\").";
      seen[key] = true;
      return "";
    });
  }

  /* ---------- deciding whether to accept a file ---------- */

  // file: {name, size}. existing: list of {fileName, fileSize} already loaded.
  // Returns {ok:true} or {ok:false, reason:"plain English"}.
  function checkIncomingFile(file, existing, config) {
    var ext = getExtension(file.name);
    if (ALLOWED_EXTENSIONS.indexOf(ext) === -1) {
      return { ok: false, reason: "Skipped \"" + file.name + "\": that file type isn't supported. Use CSV, TSV, TXT, XLS, XLSX, XLSM or ODS." };
    }
    if (file.size === 0) {
      return { ok: false, reason: "Skipped \"" + file.name + "\": the file is empty." };
    }
    if (file.size > config.MAX_FILE_MB * 1024 * 1024) {
      return { ok: false, reason: "Skipped \"" + file.name + "\": it is bigger than " + config.MAX_FILE_MB + " MB." };
    }
    for (var i = 0; i < existing.length; i++) {
      if (existing[i].fileName === file.name && existing[i].fileSize === file.size) {
        return { ok: false, reason: "Skipped \"" + file.name + "\": it is already in the list." };
      }
    }
    if (existing.length >= config.MAX_FILES) {
      return { ok: false, reason: "Skipped \"" + file.name + "\": the limit is " + config.MAX_FILES + " files at a time." };
    }
    return { ok: true };
  }

  /* ---------- reading one file ---------- */

  function startsWith(bytes, signature) {
    if (bytes.length < signature.length) return false;
    for (var i = 0; i < signature.length; i++) if (bytes[i] !== signature[i]) return false;
    return true;
  }
  var ZIP_SIGNATURE = [0x50, 0x4b];                       // "PK": xlsx, xlsm, ods
  var OLD_XLS_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0];       // classic .xls

  // Turn the raw bytes of a text file (CSV etc.) into a string.
  // Tries UTF-8 first (most common). If the bytes aren't valid UTF-8,
  // falls back to the old Windows encoding so accents still come out right.
  function decodeText(bytes) {
    var text;
    try {
      text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    } catch (e) {
      try {
        text = new TextDecoder("windows-1252").decode(bytes);
      } catch (e2) {
        text = new TextDecoder("utf-8").decode(bytes);
      }
    }
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);   // drop the invisible BOM marker
    return text;
  }

  function corruptError(fileName) {
    var err = new Error("That file looks corrupted or isn't really a spreadsheet: " + fileName);
    err.friendly = true;
    return err;
  }

  // Low-level: bytes -> SheetJS workbook.
  // keepAsText only matters for text files: it stops "001" turning into 1.
  function readWorkbook(XLSX, bytes, fileName, keepAsText) {
    var ext = getExtension(fileName);
    var isBinary = startsWith(bytes, ZIP_SIGNATURE) || startsWith(bytes, OLD_XLS_SIGNATURE);
    var isTextKind = TEXT_EXTENSIONS.indexOf(ext) !== -1 && !isBinary;

    // A file called .xlsx / .xlsm / .ods that isn't a zip inside is not a real workbook.
    if (!isBinary && (ext === "xlsx" || ext === "xlsm" || ext === "ods")) throw corruptError(fileName);

    var wb;
    try {
      if (isTextKind) {
        var options = { type: "string", raw: !!keepAsText, cellDates: false };
        if (ext === "tsv") options.FS = "\t";
        wb = XLSX.read(decodeText(bytes), options);
      } else {
        wb = XLSX.read(bytes, { type: "array", raw: !!keepAsText, cellDates: false });
      }
    } catch (e) {
      throw corruptError(fileName);
    }
    if (!wb || !wb.SheetNames || !wb.SheetNames.length) throw corruptError(fileName);
    return { workbook: wb, isTextKind: isTextKind };
  }

  // Sheet -> rows of plain text, exactly as the values are displayed.
  function sheetToTextRows(XLSX, sheet) {
    if (!sheet || !sheet["!ref"]) return [];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "", blankrows: true });
    // A file with nothing but blank lines or spaces counts as empty.
    var hasData = rows.some(function (row) {
      return row.some(function (value) { return String(value == null ? "" : value).trim() !== ""; });
    });
    return hasData ? rows : [];
  }

  /*
    Read one file. Returns an "entry":
      fileName, fileSize, kind ("text" or "workbook"),
      rows       : every value as text (used for the preview, counts and "keep as text" output)
      rowCount, colCount,
      sheetName  : the sheet that was used
      sheetCount : how many sheets the file had
      warnings   : list of plain-English notes
      sheet      : (workbook files only) the original sheet, with real numbers and dates
      bytes      : (text files only) the original bytes, so we can re-read as numbers if asked
    Throws an Error with .friendly = true if the file can't be read.
    Returns an entry with rowCount 0 if the file has nothing in it.
  */
  function parseFile(XLSX, data, fileName) {
    var bytes = toBytes(data);
    var result = readWorkbook(XLSX, bytes, fileName, true);
    var wb = result.workbook;
    var sheetName = wb.SheetNames[0];
    var sheet = wb.Sheets[sheetName];
    var rows = sheetToTextRows(XLSX, sheet);

    var colCount = 0;
    for (var r = 0; r < rows.length; r++) if (rows[r].length > colCount) colCount = rows[r].length;

    var warnings = [];
    if (wb.SheetNames.length > 1) {
      warnings.push("Only the first sheet (" + sheetName + ") was used.");
    }

    return {
      fileName: fileName,
      fileSize: bytes.length,
      kind: result.isTextKind ? "text" : "workbook",
      rows: rows,
      rowCount: rows.length,
      colCount: colCount,
      sheetName: sheetName,
      sheetCount: wb.SheetNames.length,
      warnings: warnings,
      sheet: result.isTextKind ? null : sheet,
      bytes: result.isTextKind ? bytes : null,
      tabName: ""                                          // filled in by assignTabNames or the page
    };
  }

  // Give every entry that has no tab name yet a clean, unique one.
  function assignTabNames(entries) {
    var used = {};
    entries.forEach(function (e) { if (e.tabName) used[e.tabName.toLowerCase()] = true; });
    entries.forEach(function (e) {
      if (!e.tabName) e.tabName = makeUniqueName(cleanTabName(e.fileName), used);
    });
    return entries;
  }

  /* ---------- building the combined workbook ---------- */

  // Column widths: as wide as the longest value, within sensible limits.
  function columnWidths(rows) {
    var widths = [];
    var limit = Math.min(rows.length, 2000);               // looking at 2000 rows is plenty
    for (var r = 0; r < limit; r++) {
      var row = rows[r];
      for (var c = 0; c < row.length; c++) {
        var len = String(row[c] == null ? "" : row[c]).length;
        if (!widths[c] || len > widths[c]) widths[c] = len;
      }
    }
    var out = [];
    for (var i = 0; i < widths.length; i++) {
      out.push({ wch: Math.max(MIN_COL_WIDTH, Math.min(MAX_COL_WIDTH, (widths[i] || 0) + 2)) });
    }
    return out;
  }

  // Rows of text -> a sheet where EVERY cell is stored as text.
  function textSheet(XLSX, rows) {
    var ws = {};
    var maxCol = 0;
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (row.length > maxCol) maxCol = row.length;
      for (var c = 0; c < row.length; c++) {
        var value = row[c];
        if (value === "" || value == null) continue;       // leave blanks blank
        ws[XLSX.utils.encode_cell({ r: r, c: c })] = { t: "s", v: String(value) };
      }
    }
    ws["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(rows.length - 1, 0), c: Math.max(maxCol - 1, 0) }
    });
    return ws;
  }

  /*
    entries : list from parseFile, each with a valid tabName, in the order wanted.
    options : { keepAsText: true/false, addIndex: true/false }
    Returns { workbook, tabNames (final order, including Index if added) }.
  */
  function buildWorkbook(XLSX, entries, options) {
    options = options || {};
    var keepAsText = options.keepAsText !== false;         // default ON
    if (!entries.length) throw new Error("There are no files to combine.");

    var names = entries.map(function (e) { return e.tabName; });
    var problems = validateTabNames(names).filter(Boolean);
    if (problems.length) throw new Error(problems[0]);

    var wb = XLSX.utils.book_new();
    var finalNames = [];

    if (options.addIndex) {
      var used = {};
      names.forEach(function (n) { used[n.toLowerCase()] = true; });
      var indexName = makeUniqueName("Index", used);       // becomes "Index (2)" if a tab is already "Index"
      var indexRows = [["#", "File name", "Tab name", "Rows", "Columns"]];
      entries.forEach(function (e, i) {
        indexRows.push([i + 1, e.fileName, e.tabName, e.rowCount, e.colCount]);
      });
      var indexSheet = XLSX.utils.aoa_to_sheet(indexRows);
      indexSheet["!cols"] = columnWidths(indexRows);
      XLSX.utils.book_append_sheet(wb, indexSheet, indexName);
      finalNames.push(indexName);
    }

    entries.forEach(function (e) {
      var ws;
      if (keepAsText) {
        ws = textSheet(XLSX, e.rows);
      } else if (e.kind === "text") {
        // Re-read the text file letting SheetJS spot numbers and dates.
        var again = readWorkbook(XLSX, e.bytes, e.fileName, false).workbook;
        ws = again.Sheets[again.SheetNames[0]];
      } else {
        ws = e.sheet;
      }
      ws["!cols"] = columnWidths(e.rows);
      XLSX.utils.book_append_sheet(wb, ws, e.tabName);
      finalNames.push(e.tabName);
    });

    return { workbook: wb, tabNames: finalNames };
  }

  /* ---------- writing the .xlsx bytes (with the top row frozen) ---------- */

  // The free edition of SheetJS can't write "freeze panes", so after it builds
  // the .xlsx (which is really a zip of small XML files) we open the zip and add
  // the one line of XML per sheet that tells Excel to freeze row 1.
  function freezeTopRows(XLSX, xlsxBytes) {
    var zip = XLSX.CFB.read(xlsxBytes, { type: "array" });
    var decoder = new TextDecoder("utf-8");
    var encoder = new TextEncoder();
    var pane = '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
               '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>';
    var changed = 0;
    for (var i = 0; i < zip.FullPaths.length; i++) {
      if (!/xl\/worksheets\/sheet\d+\.xml$/.test(zip.FullPaths[i])) continue;
      var entry = zip.FileIndex[i];
      var xml = decoder.decode(toBytes(entry.content));
      var patched = xml
        .replace(/<sheetView([^>]*)\/>/, "<sheetView$1>" + pane + "</sheetView>");
      if (patched !== xml) {
        entry.content = encoder.encode(patched);
        entry.size = entry.content.length;
        changed++;
      }
    }
    if (!changed) return xlsxBytes;
    return toBytes(XLSX.CFB.write(zip, { fileType: "zip", type: "array", compression: true }));
  }

  // Workbook -> bytes of a finished .xlsx file.
  function workbookToBytes(XLSX, workbook) {
    var plain = toBytes(XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true }));
    try {
      return freezeTopRows(XLSX, plain);
    } catch (e) {
      return plain;                                        // freezing is a nicety; never lose the file over it
    }
  }

  // "Combined" + a Date -> "Combined_2026-11-11.xlsx" (local date, not UTC).
  function outputFileName(prefix, date) {
    date = date || new Date();
    function two(n) { return (n < 10 ? "0" : "") + n; }
    return (prefix || "Combined") + "_" + date.getFullYear() + "-" + two(date.getMonth() + 1) + "-" + two(date.getDate()) + ".xlsx";
  }

  return {
    MAX_TAB_LEN: MAX_TAB_LEN,
    ALLOWED_EXTENSIONS: ALLOWED_EXTENSIONS,
    getExtension: getExtension,
    stripExtension: stripExtension,
    cleanTabName: cleanTabName,
    makeUniqueName: makeUniqueName,
    tabNameProblem: tabNameProblem,
    validateTabNames: validateTabNames,
    checkIncomingFile: checkIncomingFile,
    parseFile: parseFile,
    assignTabNames: assignTabNames,
    buildWorkbook: buildWorkbook,
    workbookToBytes: workbookToBytes,
    outputFileName: outputFileName
  };
});
