/* =====================================================================
   CSV Combiner: the page code.
   ---------------------------------------------------------------------
   This file handles what you SEE and CLICK: the drop zone, the file
   list, the options and the big orange button.
   The actual combining lives in js/combine-core.js (so it can be tested).
   Nothing here sends a file anywhere. It all happens in this browser tab.
   ===================================================================== */
(function () {
  "use strict";

  var CONFIG = window.COMBINER_CONFIG || { REQUIRE_LOGIN: false, MAX_FILES: 50, MAX_FILE_MB: 25, OUTPUT_PREFIX: "Combined" };
  var Core = window.CombinerCore;
  var PREVIEW_ROWS = 5;

  function $(id) { return document.getElementById(id); }

  /* ---------- login gate (switched off for the free beta) ---------- */
  if (CONFIG.REQUIRE_LOGIN) {
    // TODO: Netlify Identity check goes here.
    // When login is built: if the visitor IS logged in, skip this block
    // (don't return) so the tool shows as normal.
    $("loginGate").hidden = false;
    $("tool").hidden = true;
    return;
  }

  /* ---------- if the libraries didn't load, say so plainly ---------- */
  if (!window.XLSX || !Core) {
    $("notices").innerHTML = "";
    var broken = document.createElement("div");
    broken.className = "notice is-error";
    broken.textContent = "Part of this page didn't load, so the combiner can't run. Please refresh the page and try again.";
    $("notices").appendChild(broken);
    return;
  }

  /* ---------- the page's memory ---------- */
  var entries = [];        // one per loaded file, in tab order (see parseFile in combine-core.js)
  var notices = [];        // {text, isError} messages about skipped files
  var nextId = 1;
  var dragId = null;       // id of the row being dragged

  var els = {
    dropzone: $("dropzone"), fileInput: $("fileInput"), chooseBtn: $("chooseBtn"),
    notices: $("notices"), filesPanel: $("filesPanel"), fileList: $("fileList"),
    fileCount: $("fileCount"), clearBtn: $("clearBtn"),
    optText: $("optText"), optIndex: $("optIndex"),
    combineBtn: $("combineBtn"), goHint: $("goHint"), nameError: $("nameError"),
    result: $("result")
  };

  /* ---------- tiny DOM helpers ---------- */
  function el(tag, className, text) {
    var node = document.createElement(tag);
    if (className) node.className = className;
    if (text != null) node.textContent = text;
    return node;
  }
  function icon(pathD) {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("aria-hidden", "true");
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", pathD);
    path.setAttribute("fill", "none");
    path.setAttribute("stroke", "currentColor");
    path.setAttribute("stroke-width", "2.4");
    path.setAttribute("stroke-linecap", "round");
    path.setAttribute("stroke-linejoin", "round");
    svg.appendChild(path);
    return svg;
  }
  var ICONS = {
    grip: "M9 6h.01M15 6h.01M9 12h.01M15 12h.01M9 18h.01M15 18h.01",
    up: "M6 15l6-6 6 6",
    down: "M6 9l6 6 6-6",
    close: "M6 6l12 12M18 6L6 18"
  };
  function plural(n, word) { return n + " " + word + (n === 1 ? "" : "s"); }

  /* ---------- adding files ---------- */

  function readAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () { resolve(reader.result); };
      reader.onerror = function () { reject(reader.error || new Error("read failed")); };
      reader.readAsArrayBuffer(file);
    });
  }

  function addNotice(text, isError) { notices.push({ text: text, isError: !!isError }); }

  // Takes a FileList (or array of File). Reads them one at a time, in order.
  function addFiles(fileList) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return Promise.resolve();
    notices = [];
    hideResult();
    els.dropzone.classList.add("is-busy");

    var chain = Promise.resolve();
    files.forEach(function (file) {
      chain = chain.then(function () {
        var verdict = Core.checkIncomingFile({ name: file.name, size: file.size }, entries, CONFIG);
        if (!verdict.ok) { addNotice(verdict.reason); return; }

        return readAsArrayBuffer(file).then(function (buffer) {
          var entry = Core.parseFile(window.XLSX, buffer, file.name);
          if (entry.rowCount === 0) {
            addNotice("Skipped \"" + file.name + "\": there is no data in it.");
            return;
          }
          entry.fileSize = file.size;
          entry.id = nextId++;
          entry.previewOpen = false;
          entries.push(entry);
          Core.assignTabNames(entries);      // gives the new file a clean, unique tab name
        }).catch(function (err) {
          addNotice(err && err.friendly ? err.message
            : "That file looks corrupted or isn't really a spreadsheet: " + file.name, true);
        });
      });
    });

    return chain.catch(function () {
      addNotice("Something went wrong while reading those files. Please try again.", true);
    }).then(function () {
      els.dropzone.classList.remove("is-busy");
      els.fileInput.value = "";              // so picking the same file again still fires
      render();
    });
  }

  /* ---------- drawing the page ---------- */

  function renderNotices() {
    els.notices.innerHTML = "";
    notices.forEach(function (notice, index) {
      var box = el("div", "notice" + (notice.isError ? " is-error" : ""));
      box.appendChild(el("span", "", notice.text));
      var close = el("button", "", "×");
      close.type = "button";
      close.setAttribute("aria-label", "Dismiss this message");
      close.addEventListener("click", function () { notices.splice(index, 1); renderNotices(); });
      box.appendChild(close);
      els.notices.appendChild(box);
    });
  }

  function buildPreview(entry) {
    var wrap = el("div", "preview");
    var table = el("table");
    var shown = entry.rows.slice(0, PREVIEW_ROWS);
    var caption = el("caption", "", "First " + plural(shown.length, "row") + " of " + entry.rowCount + ".");
    table.appendChild(caption);
    var body = el("tbody");
    shown.forEach(function (row) {
      var tr = el("tr");
      for (var c = 0; c < entry.colCount; c++) {
        tr.appendChild(el("td", "", row[c] == null ? "" : String(row[c])));
      }
      body.appendChild(tr);
    });
    table.appendChild(body);
    wrap.appendChild(table);
    return wrap;
  }

  function buildRow(entry, index) {
    var row = el("li", "file-row");
    row.dataset.id = entry.id;

    /* move controls: drag handle + up/down arrows (arrows also work by keyboard and on phones) */
    var move = el("div", "row-move");
    var handle = el("span", "handle");
    handle.title = "Drag to reorder";
    handle.setAttribute("draggable", "true");
    handle.setAttribute("aria-hidden", "true");
    handle.appendChild(icon(ICONS.grip));
    handle.addEventListener("dragstart", function (ev) {
      dragId = entry.id;
      row.classList.add("is-dragging");
      ev.dataTransfer.effectAllowed = "move";
      ev.dataTransfer.setData("text/plain", String(entry.id));
      if (ev.dataTransfer.setDragImage) ev.dataTransfer.setDragImage(row, 20, 20);
    });
    handle.addEventListener("dragend", function () { dragId = null; clearDropMarks(); row.classList.remove("is-dragging"); });
    move.appendChild(handle);

    var arrows = el("div", "arrows");
    var up = el("button", "icon-btn");
    up.type = "button"; up.disabled = index === 0;
    up.setAttribute("aria-label", "Move " + entry.fileName + " up");
    up.appendChild(icon(ICONS.up));
    up.addEventListener("click", function () { moveEntry(index, index - 1, "up"); });
    var down = el("button", "icon-btn");
    down.type = "button"; down.disabled = index === entries.length - 1;
    down.setAttribute("aria-label", "Move " + entry.fileName + " down");
    down.appendChild(icon(ICONS.down));
    down.addEventListener("click", function () { moveEntry(index, index + 1, "down"); });
    arrows.appendChild(up); arrows.appendChild(down);
    move.appendChild(arrows);
    row.appendChild(move);

    /* file name + size */
    var info = el("div", "file-info");
    var name = el("div", "file-name", entry.fileName);
    name.title = entry.fileName;
    info.appendChild(name);
    info.appendChild(el("div", "file-meta", plural(entry.rowCount, "row") + " × " + plural(entry.colCount, "column")));
    row.appendChild(info);

    /* editable tab name */
    var field = el("div", "tab-field");
    var inputId = "tab-" + entry.id;
    var label = el("label", "", "Tab name");
    label.setAttribute("for", inputId);
    var input = el("input");
    input.type = "text"; input.id = inputId; input.value = entry.tabName;
    input.maxLength = 60;                    // lets you paste long names; we flag anything over 31
    input.autocomplete = "off"; input.spellcheck = false;
    input.setAttribute("aria-describedby", "err-" + entry.id);
    input.addEventListener("input", function () { entry.tabName = input.value; validate(); });
    field.appendChild(label); field.appendChild(input);
    row.appendChild(field);

    /* preview + remove */
    var actions = el("div", "row-actions");
    var previewBtn = el("button", "preview-btn", "Preview");
    previewBtn.type = "button";
    previewBtn.setAttribute("aria-expanded", entry.previewOpen ? "true" : "false");
    previewBtn.setAttribute("aria-label", "Preview the first rows of " + entry.fileName);
    previewBtn.addEventListener("click", function () { entry.previewOpen = !entry.previewOpen; render(); });
    var remove = el("button", "icon-btn remove");
    remove.type = "button";
    remove.setAttribute("aria-label", "Remove " + entry.fileName);
    remove.appendChild(icon(ICONS.close));
    remove.addEventListener("click", function () {
      entries.splice(entries.indexOf(entry), 1);
      hideResult();
      render();
    });
    actions.appendChild(previewBtn); actions.appendChild(remove);
    row.appendChild(actions);

    /* warnings, name error, preview table */
    var extra = el("div", "row-extra");
    entry.warnings.forEach(function (w) { extra.appendChild(el("div", "row-warning", "⚠ " + w)); });
    var error = el("div", "row-error");
    error.id = "err-" + entry.id;
    error.hidden = true;
    extra.appendChild(error);
    if (entry.previewOpen) extra.appendChild(buildPreview(entry));
    row.appendChild(extra);

    /* this row as a drop target for reordering */
    row.addEventListener("dragover", function (ev) {
      if (dragId == null || dragId === entry.id) return;
      ev.preventDefault();
      ev.stopPropagation();
      var box = row.getBoundingClientRect();
      var after = ev.clientY > box.top + box.height / 2;
      clearDropMarks();
      row.classList.add(after ? "drop-after" : "drop-before");
    });
    row.addEventListener("drop", function (ev) {
      if (dragId == null) return;
      ev.preventDefault();
      ev.stopPropagation();
      var from = indexOfId(dragId);
      var to = indexOfId(entry.id);
      var after = row.classList.contains("drop-after");
      dragId = null;
      if (from === -1 || to === -1 || from === to) { clearDropMarks(); return; }
      var moved = entries.splice(from, 1)[0];
      to = indexOfId(entry.id) + (after ? 1 : 0);
      entries.splice(to, 0, moved);
      hideResult();
      render();
    });

    return row;
  }

  function indexOfId(id) {
    for (var i = 0; i < entries.length; i++) if (entries[i].id === id) return i;
    return -1;
  }
  function clearDropMarks() {
    var rows = els.fileList.children;
    for (var i = 0; i < rows.length; i++) rows[i].classList.remove("drop-before", "drop-after");
  }
  function moveEntry(from, to, direction) {
    if (to < 0 || to >= entries.length) return;
    var moved = entries.splice(from, 1)[0];
    entries.splice(to, 0, moved);
    hideResult();
    render();
    // keep keyboard focus on the same file's arrow so you can keep pressing it
    var row = els.fileList.children[to];
    if (row) {
      var buttons = row.querySelectorAll(".arrows .icon-btn");
      var target = direction === "up" ? buttons[0] : buttons[1];
      if (target && target.disabled) target = direction === "up" ? buttons[1] : buttons[0];
      if (target) target.focus();
    }
  }

  // Check the tab names and switch the big button on or off. Runs on every keystroke.
  function validate() {
    var problems = Core.validateTabNames(entries.map(function (e) { return e.tabName; }));
    var firstProblem = "";
    entries.forEach(function (entry, i) {
      var row = els.fileList.children[i];
      if (!row) return;
      var input = row.querySelector(".tab-field input");
      var error = row.querySelector(".row-error");
      var bad = !!problems[i];
      row.classList.toggle("has-error", bad);
      input.setAttribute("aria-invalid", bad ? "true" : "false");
      error.hidden = !bad;
      error.textContent = bad ? problems[i] : "";
      if (bad && !firstProblem) firstProblem = "Fix the tab name for \"" + entry.fileName + "\": " + problems[i];
    });

    els.nameError.hidden = !firstProblem;
    els.nameError.textContent = firstProblem;
    els.combineBtn.disabled = entries.length === 0 || !!firstProblem;
    els.goHint.textContent = entries.length === 0
      ? "Add at least one file to get started."
      : firstProblem
        ? "The download is paused until the tab name above is fixed."
        : plural(entries.length, "file") + " ready. You'll get one workbook with " +
          plural(entries.length + (els.optIndex.checked ? 1 : 0), "tab") + ".";
    return !firstProblem;
  }

  function render() {
    renderNotices();
    els.filesPanel.hidden = entries.length === 0;
    els.fileCount.textContent = entries.length ? "(" + entries.length + ")" : "";
    els.fileList.innerHTML = "";
    entries.forEach(function (entry, index) { els.fileList.appendChild(buildRow(entry, index)); });
    validate();
  }

  /* ---------- result message ---------- */

  function hideResult() { els.result.hidden = true; els.result.innerHTML = ""; }

  function showResult(mainText, warnings, isError) {
    els.result.innerHTML = "";
    els.result.classList.toggle("is-error", !!isError);
    els.result.appendChild(el("p", "result-main", mainText));
    if (warnings && warnings.length) {
      var list = el("ul");
      warnings.forEach(function (w) { list.appendChild(el("li", "", w)); });
      els.result.appendChild(list);
    }
    els.result.hidden = false;
  }

  /* ---------- combine and download ---------- */

  function downloadBytes(bytes, fileName) {
    var blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  }

  function combine() {
    if (!validate()) return;
    var originalLabel = els.combineBtn.textContent;
    els.combineBtn.disabled = true;
    els.combineBtn.textContent = "Combining...";

    // Small pause so the button visibly changes before the heavy work starts.
    setTimeout(function () {
      try {
        var built = Core.buildWorkbook(window.XLSX, entries, {
          keepAsText: els.optText.checked,
          addIndex: els.optIndex.checked
        });
        var bytes = Core.workbookToBytes(window.XLSX, built.workbook);
        var fileName = Core.outputFileName(CONFIG.OUTPUT_PREFIX, new Date());
        downloadBytes(bytes, fileName);

        var warnings = [];
        entries.forEach(function (e) {
          e.warnings.forEach(function (w) { warnings.push(e.fileName + ": " + w); });
        });
        notices.forEach(function (n) { warnings.push(n.text); });
        showResult("Done. " + plural(entries.length, "file") + " combined into " + fileName, warnings, false);
      } catch (err) {
        showResult("Sorry, the workbook couldn't be built. " + (err && err.message ? err.message : "Please try again."), [], true);
      }
      els.combineBtn.textContent = originalLabel;
      validate();
    }, 30);
  }

  /* ---------- wiring up the buttons ---------- */

  els.chooseBtn.addEventListener("click", function () { els.fileInput.click(); });
  els.fileInput.addEventListener("change", function () { addFiles(els.fileInput.files); });
  els.clearBtn.addEventListener("click", function () { entries = []; notices = []; hideResult(); render(); });
  els.combineBtn.addEventListener("click", combine);
  els.optIndex.addEventListener("change", function () { hideResult(); validate(); });
  els.optText.addEventListener("change", hideResult);

  // Dropping files: the whole page accepts them, so a near miss doesn't make
  // the browser navigate away and open the file instead.
  function isFileDrag(ev) {
    var types = ev.dataTransfer && ev.dataTransfer.types;
    if (!types) return false;
    for (var i = 0; i < types.length; i++) if (types[i] === "Files") return true;
    return false;
  }
  document.addEventListener("dragover", function (ev) {
    if (!isFileDrag(ev)) return;
    ev.preventDefault();
    els.dropzone.classList.add("is-over");
  });
  document.addEventListener("dragleave", function (ev) {
    if (!ev.relatedTarget) els.dropzone.classList.remove("is-over");
  });
  document.addEventListener("drop", function (ev) {
    if (!isFileDrag(ev)) return;
    ev.preventDefault();
    els.dropzone.classList.remove("is-over");
    addFiles(ev.dataTransfer.files);
  });

  // Handy for testing from the browser console; not used by the page itself.
  window.CombinerApp = { addFiles: addFiles, getEntries: function () { return entries; } };

  render();
})();
