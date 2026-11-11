/* =====================================================================
   Makes the sample files the tests (and you) can try the app with.
   Run:  node tests/make-samples.js
   Output goes to tests/output/samples/ (ignored by git).

   One wrinkle: Windows does not allow ":" in a file name, so the sample
   whose real name is  big_names_and_bad_characters [test]:1.csv  is saved
   on disk with "_" instead of ":". The file tests/output/samples/manifest.json
   remembers the intended name of every sample, and the tests use that.
   ===================================================================== */
"use strict";
const fs = require("node:fs");
const path = require("node:path");
const XLSX = require("../vendor/xlsx.full.min.js");

const OUT_DIR = path.join(__dirname, "output", "samples");

// Each sample: the name it is MEANT to have, and its contents.
function buildSamples() {
  const samples = [];
  const text = (name, content) => samples.push({ name, bytes: Buffer.from(content, "utf8") });

  // Trimble-style points: point number, easting, northing, elevation, description.
  text("points_a.csv", [
    "Point,Easting,Northing,Elevation,Description",
    "001,6254871.1250,1847732.0400,512.340,CP NAIL",
    "002,6254902.8810,1847755.6120,512.910,CP NAIL",
    "0042,6255010.0000,1847801.5000,513.000,GRID A-1",
    "100,6255044.2500,1847833.7500,513.250,GRID A-2",
    "00123,6255100.1234567,1847900.7654321,514.005,COL FTG"
  ].join("\r\n") + "\r\n");

  text("points_b.csv", [
    "Point,Easting,Northing,Elevation,Description",
    "0501,6254000.500,1847000.250,498.120,SLEEVE 4IN",
    "0502,6254003.500,1847000.250,498.120,SLEEVE 4IN",
    "0503,6254006.500,1847000.250,498.120,HANGER",
    "1/2,6254009.500,1847000.250,498.120,LOOKS LIKE A DATE"
  ].join("\n") + "\n");

  text("tab_separated.tsv", [
    "Point\tEasting\tNorthing\tElevation\tDescription",
    "007\t1000.000\t2000.000\t10.50\tBM, BRASS CAP",
    "008\t1010.000\t2010.000\t10.75\tBM"
  ].join("\n") + "\n");

  // Tests tab-name cleaning: too long, and full of characters Excel forbids.
  text("big_names_and_bad_characters [test]:1.csv", "Point,Note\n01,long name\n02,bad characters\n");

  // Two files that would both become the tab "duplicate name".
  text("duplicate name.csv", "Point,Value\n01,first file\n");
  text("duplicate name.tsv", "Point\tValue\n01\tsecond file\n02\tsecond file again\n");

  // 0 bytes.
  samples.push({ name: "empty.csv", bytes: Buffer.alloc(0) });

  // A real workbook with two sheets (only the first should be used).
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Point", "Easting", "Northing"],
    ["A1", 100.5, 200.25],
    ["A2", 110.5, 210.25],
    ["A3", 120.5, 220.25]
  ]), "FirstSheet");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["This sheet", "should be", "ignored"],
    ["x", "y", "z"]
  ]), "SecondSheet");
  samples.push({ name: "real.xlsx", bytes: Buffer.from(XLSX.write(wb, { bookType: "xlsx", type: "buffer" })) });

  // Quoted commas, a doubled quote, an accent, and a UTF-8 BOM at the front.
  text("quoted_commas_bom.csv", "﻿" + [
    "Point,Description,Crew",
    "0001,\"WALL, NORTH FACE\",José",
    "0002,\"ANCHOR 3/4\"\" DIA, EPOXY\",Zoë",
    "0003,PLAIN,Bob"
  ].join("\r\n") + "\r\n");

  return samples;
}

// Windows-safe version of a name for saving to disk.
function diskName(name) {
  return name.replace(/[:*?"<>|]/g, "_");
}

function writeSamples() {
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const manifest = [];
  for (const sample of buildSamples()) {
    const onDisk = diskName(sample.name);
    fs.writeFileSync(path.join(OUT_DIR, onDisk), sample.bytes);
    manifest.push({ name: sample.name, file: onDisk, size: sample.bytes.length });
  }
  fs.writeFileSync(path.join(OUT_DIR, "manifest.json"), JSON.stringify(manifest, null, 2));
  return manifest;
}

module.exports = { buildSamples, writeSamples, OUT_DIR };

if (require.main === module) {
  const manifest = writeSamples();
  console.log("Wrote " + manifest.length + " sample files to " + OUT_DIR);
}
