/*
 * Mantis Updater - Office.js version of "Mantis Updater V1.0.0.xlam".
 * Mantis Networks (Pty) Ltd
 *
 * VBA module          ->  here
 *   modUtil           ->  "helpers" + "Excel helpers"
 *   modPatch          ->  applyPatchTable / applyOnePatch
 *   modAIGuide        ->  ensureAISheets / buildGuideSheet / buildAISheet
 *   modAI             ->  runQuestion / callProvider / exportSnapshot / clearAI
 *   modMain           ->  applyUpdates / fetchPatches / writeVersionsLog
 *   ThisWorkbook menu ->  ribbon button in manifest.xml + taskpane.html
 *
 * Not possible in Office.js (see README.md):
 *   - VBModule / VBRemove patch rows (no access to the VBA project) -> skipped and logged
 *   - form buttons on the AI sheet -> the buttons live in the task pane instead
 *   - SaveCopyAs backup (was switched off in the prototype anyway)
 */

/* global Excel, Office, XLSX */

var MU = (function () {
  "use strict";

  var ADDIN_NAME = "Mantis Updater";
  var ADDIN_VERSION = "1.0.0";
  var GUIDE_SHEET = "AI Guide";
  var AI_SHEET = "AI";
  var APPLIED_PREFIX = "MantisPatch_";
  var MANTIS_BLUE = "#4682B4";
  var LIGHT_FILL = "#F2F2F2";
  var INPUT_FILL = "#FFFFE6";
  var MAX_SCORECARD_ROWS = 62;
  var ANSWER_PLACEHOLDER = "(the answer appears here)";

  var PATCH_COLUMNS = ["PatchID", "Product", "MinVer", "MaxVer", "Type", "Sheet", "Range",
    "Find", "Replace", "FillHex", "FontHex", "Expect", "Note"];

  var DEFAULT_CONFIG = {
    PatchURL: "",
    MantisAIEndpoint: "",
    AnthropicModel: "claude-sonnet-4-5",
    OpenAIModel: "gpt-4o-mini"
  };

  // =====================================================================================
  // Storage (replaces the Patches / Config / Log sheets of the .xlam)
  // =====================================================================================
  var memoryStore = {};
  function storeGet(key, def) {
    try {
      var v = window.localStorage.getItem("mantisUpdater." + key);
      return v == null ? def : JSON.parse(v);
    } catch (e) {
      return key in memoryStore ? memoryStore[key] : def;
    }
  }
  function storeSet(key, value) {
    memoryStore[key] = value;
    try { window.localStorage.setItem("mantisUpdater." + key, JSON.stringify(value)); } catch (e) { /* private mode */ }
  }

  function getConfig() {
    var c = storeGet("config", {});
    var out = {};
    Object.keys(DEFAULT_CONFIG).forEach(function (k) { out[k] = c[k] != null ? c[k] : DEFAULT_CONFIG[k]; });
    return out;
  }
  function setConfig(c) { storeSet("config", c); }
  function cfgValue(key) { return String(getConfig()[key] || "").trim(); }

  function getPatches() { return storeGet("patches", []); }
  function setPatches(rows) { storeSet("patches", rows); }

  var logListeners = [];
  function getLog() { return storeGet("log", []); }
  function clearLog() { storeSet("log", []); logListeners.forEach(function (f) { f(); }); }
  function onLog(fn) { logListeners.push(fn); }
  function logLine(msg) {
    var log = getLog();
    log.push({ t: stamp(new Date(), true), msg: String(msg) });
    if (log.length > 2000) log = log.slice(log.length - 2000);
    storeSet("log", log);
    if (typeof console !== "undefined") console.log("[Mantis Updater] " + msg);
    logListeners.forEach(function (f) { f(); });
  }

  // =====================================================================================
  // Helpers (pure - tested in tests.js)
  // =====================================================================================
  function trim(v) { return v == null ? "" : String(v).trim(); }
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function stamp(d, withSeconds) {
    return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) + " " +
      pad(d.getHours()) + ":" + pad(d.getMinutes()) + (withSeconds ? ":" + pad(d.getSeconds()) : "");
  }
  function fileStamp(d) {
    return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + "-" + pad(d.getHours()) + pad(d.getMinutes());
  }
  var MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  function longDate(d) { return pad(d.getDate()) + " " + MONTHS[d.getMonth()] + " " + d.getFullYear(); }
  function excelDateSerial(d) {
    return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  // Compare "V15.10.0" style versions: returns -1, 0, 1
  function compareVersion(a, b) {
    var pa = String(a || "").trim().toUpperCase().replace(/V/g, "").split(".");
    var pb = String(b || "").trim().toUpperCase().replace(/V/g, "").split(".");
    for (var i = 0; i < 4; i++) {
      var va = /^\s*\d+\s*$/.test(pa[i] || "") ? parseInt(pa[i], 10) : 0;
      var vb = /^\s*\d+\s*$/.test(pb[i] || "") ? parseInt(pb[i], 10) : 0;
      if (va < vb) return -1;
      if (va > vb) return 1;
    }
    return 0;
  }

  // "RRGGBB" / "#RRGGBB" / "0xRRGGBB" -> "#RRGGBB", or null if not valid
  function normHex(h) {
    var s = trim(h).replace(/^#/, "").replace(/^0x/i, "");
    return /^[0-9a-f]{6}$/i.test(s) ? "#" + s.toUpperCase() : null;
  }

  function escapeRegExp(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

  // Find/replace that will not run twice: if the cell already contains the replacement, leave it.
  function replaceText(curText, findTxt, repl) {
    if (!findTxt) return repl;
    if (repl && curText.toLowerCase().indexOf(repl.toLowerCase()) >= 0) return curText;
    return curText.replace(new RegExp(escapeRegExp(findTxt), "gi"), function () { return repl; });
  }

  // Split one CSV line, honouring double quotes
  function splitCsvLine(line) {
    var out = [], cur = "", inQ = false;
    for (var i = 0; i < line.length; i++) {
      var c = line.charAt(i);
      if (c === '"') {
        if (inQ && line.charAt(i + 1) === '"') { cur += '"'; i++; } else { inQ = !inQ; }
      } else if (c === "," && !inQ) {
        out.push(cur); cur = "";
      } else {
        cur += c;
      }
    }
    out.push(cur);
    return out;
  }

  // CSV text (header row + 13 columns, same as the Patches sheet) -> array of patch objects
  function parsePatchCsv(text) {
    var lines = String(text || "").replace(/^﻿/, "").replace(/\r/g, "").split("\n");
    var rows = [];
    for (var i = 1; i < lines.length; i++) {
      if (!lines[i].trim()) continue;
      var f = splitCsvLine(lines[i]);
      var row = {};
      PATCH_COLUMNS.forEach(function (col, j) { row[col] = j < f.length ? f[j].trim() : ""; });
      rows.push(row);
    }
    return rows;
  }

  function patchesToCsv(rows) {
    function q(v) { v = v == null ? "" : String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }
    return [PATCH_COLUMNS.join(",")].concat(rows.map(function (r) {
      return PATCH_COLUMNS.map(function (c) { return q(r[c]); }).join(",");
    })).join("\r\n");
  }

  function patchApplies(p, product, ver) {
    var prod = trim(p.Product), mn = trim(p.MinVer), mx = trim(p.MaxVer);
    if (prod !== "*" && prod.toLowerCase() !== String(product || "").toLowerCase()) return false;
    if (mn && compareVersion(ver, mn) < 0) return false;
    if (mx && compareVersion(ver, mx) > 0) return false;
    return true;
  }

  function textOf(v) {
    if (v === true) return "True";
    if (v === false) return "False";
    return v == null ? "" : String(v);
  }

  function round2(v) {
    if (typeof v === "number") return String(Math.round(v * 100) / 100);
    if (typeof v === "string" && v.trim() !== "" && !isNaN(Number(v))) return String(Math.round(Number(v) * 100) / 100);
    return textOf(v);
  }

  function colLetter(c0) {
    var s = "", n = c0 + 1;
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function cellAddress(r0, c0) { return colLetter(c0) + (r0 + 1); }

  // Question category for the usage log - never the question text
  function questionCategory(q) {
    q = String(q || "").toLowerCase();
    if (q.indexOf("next level") >= 0 || q.indexOf("improve") >= 0) return "next-level";
    if (q.indexOf("explain") >= 0) return "explain";
    if (q.indexOf("skills") >= 0 || q.indexOf("ownership") >= 0 || q.indexOf("supplier") >= 0) return "element";
    return "other";
  }

  // Pull the answer out of an Anthropic, OpenAI or Mantis-proxy response
  function extractAnswer(text) {
    var j = null;
    try { j = JSON.parse(text); } catch (e) { /* not JSON */ }
    if (j) {
      if (Array.isArray(j.content)) {
        var parts = j.content.filter(function (b) { return b && b.type === "text"; }).map(function (b) { return b.text; });
        if (parts.length) return parts.join("\n");
      }
      if (j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content) return String(j.choices[0].message.content);
      if (typeof j.answer === "string") return j.answer;
      if (typeof j.text === "string") return j.text;
      if (typeof j.content === "string") return j.content;
    }
    var m = /"(?:text|content)"\s*:\s*"((?:[^"\\]|\\.)*)"/.exec(String(text || ""));
    if (m) { try { return JSON.parse('"' + m[1] + '"'); } catch (e) { return m[1]; } }
    return "";
  }

  // CHAR-based column width (VBA ColumnWidth) -> points (Office.js columnWidth)
  function charsToPoints(chars) { return Math.round(chars * 5.25 + 3.75); }

  // =====================================================================================
  // Excel helpers
  // =====================================================================================
  async function sheetOrNull(ctx, name) {
    var ws = ctx.workbook.worksheets.getItemOrNullObject(name);
    await ctx.sync();
    return ws.isNullObject ? null : ws;
  }

  // Range behind a defined name (sheet-scoped on `scopeSheet` first, then workbook), or null
  async function namedRange(ctx, nm, scopeSheet) {
    var candidates = [];
    if (scopeSheet) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject(scopeSheet);
      await ctx.sync();
      if (!ws.isNullObject) candidates.push(ws.names.getItemOrNullObject(nm));
    }
    candidates.push(ctx.workbook.names.getItemOrNullObject(nm));
    await ctx.sync();
    for (var i = 0; i < candidates.length; i++) {
      if (candidates[i].isNullObject) continue;
      var r = candidates[i].getRangeOrNullObject();
      await ctx.sync();
      if (!r.isNullObject) return r;
    }
    return null;
  }

  async function nameExists(ctx, nm) {
    var n = ctx.workbook.names.getItemOrNullObject(nm);
    await ctx.sync();
    return !n.isNullObject;
  }

  // first cell of a named range, as text ("" if missing / error); numbers rounded to 2 when round=true
  async function nameValue(ctx, nm, round, scopeSheet) {
    try {
      var r = await namedRange(ctx, nm, scopeSheet);
      if (!r) return "";
      var c = r.getCell(0, 0);
      c.load("values,valueTypes");
      await ctx.sync();
      if (c.valueTypes[0][0] === "Error") return "";
      return round ? round2(c.values[0][0]) : textOf(c.values[0][0]);
    } catch (e) {
      return "";
    }
  }

  async function setNameValue(ctx, nm, value) {
    var r = await namedRange(ctx, nm);
    if (!r) throw new Error("Name " + nm + " not found in this calculator.");
    r.getCell(0, 0).values = [[value]];
    await ctx.sync();
  }

  async function setName(ctx, nm, range) {
    var n = ctx.workbook.names.getItemOrNullObject(nm);
    await ctx.sync();
    if (!n.isNullObject) { n.delete(); await ctx.sync(); }
    ctx.workbook.names.add(nm, range);
    await ctx.sync();
  }

  // What the VBA got from IsMantisCalc / CalcShortName / CalcVersion / CalcPassword
  async function calcInfo(ctx) {
    var wb = ctx.workbook;
    wb.load("name");
    var info = { workbookName: "", isCalc: false, product: "", version: "", password: "", hasAISheet: false };
    try { await ctx.sync(); info.workbookName = wb.name; } catch (e) { /* older hosts */ }

    var versions = await sheetOrNull(ctx, "Versions");
    var about = await sheetOrNull(ctx, "About");
    info.isCalc = !!(versions && about && (await nameExists(ctx, "Version_Num")) && (await nameExists(ctx, "Totalscore")));
    if (!info.isCalc) return info;

    info.product = await nameValue(ctx, "Short_Name", false, "About");
    if (!info.product) {
      var c5 = about.getRange("C5"); c5.load("values"); await ctx.sync();
      info.product = textOf(c5.values[0][0]);
    }
    info.version = await nameValue(ctx, "Version_Num");
    try {
      var dash = await sheetOrNull(ctx, "Dashboard");
      if (dash) { var iv = dash.getRange("IV1"); iv.load("values"); await ctx.sync(); info.password = textOf(iv.values[0][0]); }
    } catch (e) { /* no password cell */ }
    info.hasAISheet = await nameExists(ctx, "AI_Question");
    return info;
  }

  async function unprotectSheetSafe(ctx, ws, pwd) {
    ws.protection.load("protected");
    await ctx.sync();
    if (!ws.protection.protected) return true;
    try { ws.protection.unprotect(pwd || undefined); await ctx.sync(); } catch (e) {
      try { ws.protection.unprotect(); await ctx.sync(); } catch (e2) { /* still locked */ }
    }
    ws.protection.load("protected");
    await ctx.sync();
    return !ws.protection.protected;
  }

  async function protectSheetSafe(ctx, ws, pwd) {
    try {
      ws.protection.protect({ allowFormatCells: true }, pwd || undefined);
      await ctx.sync();
    } catch (e) { /* leave as is */ }
  }

  async function unprotectBookSafe(ctx, pwd) {
    var p = ctx.workbook.protection;
    p.load("protected");
    await ctx.sync();
    if (!p.protected) return true;
    try { p.unprotect(pwd || undefined); await ctx.sync(); } catch (e) {
      try { p.unprotect(); await ctx.sync(); } catch (e2) { /* still locked */ }
    }
    p.load("protected");
    await ctx.sync();
    return !p.protected;
  }

  async function protectBookSafe(ctx, pwd) {
    try { ctx.workbook.protection.protect(pwd || undefined); await ctx.sync(); } catch (e) { /* leave */ }
  }

  async function saveWorkbook(ctx) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.11")) {
      logLine("   save: this Excel cannot save from an add-in - press Ctrl+S");
      return false;
    }
    try {
      ctx.workbook.save(Excel.SaveBehavior.save);
      await ctx.sync();
      return true;
    } catch (e) {
      logLine("   WARNING: could not save (" + e.message + ") - press Ctrl+S");
      return false;
    }
  }

  // =====================================================================================
  // modPatch - data-driven patching
  // =====================================================================================
  async function patchAlreadyApplied(ctx, id) {
    var p = ctx.workbook.properties.custom.getItemOrNullObject(APPLIED_PREFIX + id);
    await ctx.sync();
    return !p.isNullObject;
  }

  async function markPatchApplied(ctx, id) {
    try {
      ctx.workbook.properties.custom.add(APPLIED_PREFIX + id, stamp(new Date(), false));
      await ctx.sync();
    } catch (e) {
      logLine("   WARNING: could not record " + id + " as applied (" + e.message + ")");
    }
  }

  async function applyPatchTable(ctx, info, pwd) {
    var rows = getPatches();
    var applied = 0, skipped = 0, notes = "";
    for (var i = 0; i < rows.length; i++) {
      var p = rows[i];
      var id = trim(p.PatchID);
      if (!id || !patchApplies(p, info.product, info.version)) continue;
      var res;
      if (await patchAlreadyApplied(ctx, id)) {
        res = "SKIP already applied to this calculator";
      } else {
        try { res = await applyOnePatch(ctx, p, pwd); } catch (e) { res = "ERR " + e.message; }
      }
      if (res.indexOf("OK") === 0) {
        applied++;
        await markPatchApplied(ctx, id);
        notes += id + " " + trim(p.Note) + "; ";
      } else {
        skipped++;
      }
      logLine(id + " -> " + res);
    }
    return { applied: applied, skipped: skipped, notes: notes };
  }

  // One patch row. Returns "OK ..." or "SKIP ..." / "ERR ..."
  async function applyOnePatch(ctx, p, pwd) {
    var typ = trim(p.Type).toLowerCase();
    if (typ === "vbmodule" || typ === "vbremove") {
      return "SKIP " + p.Type + " changes VBA code, which an Office add-in cannot do - use the VBA updater (.xlam) for this row";
    }
    if (["format", "rename", "value", "formula"].indexOf(typ) < 0) return "SKIP unknown type '" + typ + "'";

    var shName = trim(p.Sheet), rngTxt = trim(p.Range);
    if (!rngTxt) return "SKIP no range given";
    var ws = await sheetOrNull(ctx, shName);
    if (!ws) return "SKIP sheet not found: " + shName;

    ws.protection.load("protected");
    await ctx.sync();
    var wasProtected = ws.protection.protected;
    if (!(await unprotectSheetSafe(ctx, ws, pwd))) return "SKIP sheet is protected and password did not work: " + shName;

    var repl = p.Replace == null ? "" : String(p.Replace);
    if (typ === "formula" && repl.charAt(0) !== "=") repl = "=" + repl;
    var findTxt = p.Find == null ? "" : String(p.Find);
    var expect = p.Expect == null ? "" : String(p.Expect);
    var fill = normHex(p.FillHex), font = normHex(p.FontHex);
    var cnt = 0;

    try {
      var parts = rngTxt.split(",");
      for (var k = 0; k < parts.length; k++) {
        var rg = ws.getRange(parts[k].trim());
        rg.load("values,formulas,valueTypes,rowCount,columnCount,rowIndex,columnIndex");
        await ctx.sync();
        for (var i = 0; i < rg.rowCount; i++) {
          for (var j = 0; j < rg.columnCount; j++) {
            var addr = shName + "!" + cellAddress(rg.rowIndex + i, rg.columnIndex + j);
            if (rg.valueTypes[i][j] === "Error") { logLine("   " + addr + " shows an error value - left alone"); continue; }
            var cur = textOf(rg.values[i][j]);

            if (expect && cur.toLowerCase() !== expect.toLowerCase() && cur.toLowerCase().indexOf(expect.toLowerCase()) < 0) {
              logLine("   guard: " + addr + " is '" + cur.slice(0, 40) + "', expected '" + expect + "' - left alone");
              continue;
            }

            var f = rg.formulas[i][j];
            var hasFormula = typeof f === "string" && f.charAt(0) === "=";
            var cell = rg.getCell(i, j);
            var changed = false, newText;

            if (typ === "format") {
              if (fill) { cell.format.fill.color = fill; changed = true; }
              if (font) { cell.format.font.color = font; changed = true; }
              if (repl) {
                if (hasFormula) logLine("   " + addr + " has a formula - text not changed");
                else { newText = replaceText(cur, findTxt, repl); if (newText !== cur) { cell.values = [[newText]]; changed = true; } }
              }
            } else if (typ === "rename") {
              if (hasFormula) logLine("   " + addr + " has a formula - left alone");
              else { newText = replaceText(cur, findTxt, repl); if (newText !== cur) { cell.values = [[newText]]; changed = true; } }
            } else if (typ === "value") {
              cell.values = [[repl]]; changed = true;
            } else if (typ === "formula") {
              cell.formulas = [[repl]]; changed = true;
            }
            if (changed) cnt++;
          }
        }
        await ctx.sync();
      }
    } finally {
      if (wasProtected) await protectSheetSafe(ctx, ws, pwd);
    }

    return cnt === 0 ? "SKIP no cells changed on " + shName + " (guard / already done / formulas)"
      : "OK " + cnt + " cell(s) on " + shName;
  }

  // =====================================================================================
  // modAIGuide - AI Guide + AI sheets
  // =====================================================================================
  async function ensureAISheets(ctx, info) {
    var msg = "";
    var guide = await sheetOrNull(ctx, GUIDE_SHEET);
    if (!guide) { await buildGuideSheet(ctx, info); msg += "AI Guide added; "; }
    else { await refreshGuideText(ctx, guide, info); msg += "AI Guide refreshed; "; }
    var ai = await sheetOrNull(ctx, AI_SHEET);
    if (!ai) { await buildAISheet(ctx); msg += "AI sheet added; "; }
    else msg += "AI sheet already there; ";
    return msg;
  }

  async function addSheetAfter(ctx, name, afterName) {
    var after = await sheetOrNull(ctx, afterName);
    var ws = ctx.workbook.worksheets.add(name);
    if (after) { after.load("position"); await ctx.sync(); ws.position = after.position + 1; }
    ws.tabColor = MANTIS_BLUE;
    await ctx.sync();
    return ws;
  }

  async function buildGuideSheet(ctx, info) {
    var ws = await addSheetAfter(ctx, GUIDE_SHEET, "About");
    ws.getRange("A:A").format.columnWidth = charsToPoints(3);
    ws.getRange("B:B").format.columnWidth = charsToPoints(120);
    await ctx.sync();
    await refreshGuideText(ctx, ws, info);
  }

  async function refreshGuideText(ctx, ws, info) {
    var nv = {};
    var names = ["OWeight", "MCWeight", "SDWeight", "ESDWeight", "SEDWeight"];
    for (var i = 0; i < names.length; i++) nv[names[i]] = await nameValue(ctx, names[i]);
    var product = info.product, ver = info.version;

    var L = [];
    function W(txt, sz, bold) { L.push({ t: txt, sz: sz, b: bold }); }
    function gap() { L.push(null); }

    W("AI Guide for this calculator", 16, true);
    W("Mantis " + product + " " + ver + "  -  this tab is written for people AND for AI assistants (Claude, Copilot, ChatGPT).", 10, false);
    gap();
    W("FOR PEOPLE - three ways to use AI with this calculator", 12, true);
    W("1. Ask AI (Mantis AI Tier or your own key): open Mantis Updater on the Home tab, go to AI, type a question and click Ask AI. Only your question and the scorecard values are sent.", 10, false);
    W("2. Claude in Excel / Copilot: open the pane and ask. Tell it to read this 'AI Guide' sheet first so it knows where the numbers are.", 10, false);
    W("3. Claude.ai / ChatGPT on the web: click 'Export AI snapshot' in the Mantis Updater pane. It makes a small values-only file you can upload. Never upload the calculator itself.", 10, false);
    gap();
    W("FOR AI ASSISTANTS - read this before answering", 12, true);
    W("You are reading a Mantis " + product + " B-BBEE calculator (Amended Codes of Good Practice, South Africa). The authoritative results are on the 'Scorecard' sheet. Your job is to READ and EXPLAIN. Never recompute, estimate, round, or modify values. If asked to change anything, refuse and point the user to the input sheets.", 10, false);
    gap();
    W("WHERE THE NUMBERS ARE (named ranges)", 12, true);
    W("Totalscore = total B-BBEE points (Scorecard!I57).  BEELevel = recognition level as scored (Scorecard!I62).  BEELevel_After = level after discounting/YES (Scorecard!J62).", 10, false);
    W("OwnScore, MCScore, SDScore, ESDScore, SEDScore = points per element.  OWeight, MCWeight, SDWeight, ESDWeight, SEDWeight = maximum points per element.", 10, false);
    W("Scorecard rows 11-56: column B indicator, C-G target/actual detail from the element sheets, H points available, I points scored.  Company_Name = Dashboard!B64.", 10, false);
    W("Element detail sheets: 'OWN Element', 'MC Element', 'SD Element', 'ESD Element', 'SED Element'. Input sheets: Ownership, MC, Employee List, SD Spend, SD Bursaries, SD Learners, ED and SD, Suppliers, SED, YES.", 10, false);
    W("Planning sheets (if present): 'Target vs Actual' (shortfall per indicator) and 'SD Scenario' (minimum learners and spend to hold or lift a level).", 10, false);
    gap();
    W("WHAT THE ELEMENTS MEAN (current maximum points in this file)", 12, true);
    W("Ownership (" + nv.OWeight + " pts): black voting rights, economic interest, net value (the 'net value' target is a priority sub-minimum).", 10, false);
    W("Management Control (" + nv.MCWeight + " pts): black and black-female representation at board, executive, senior, middle and junior management, plus disabled employees.", 10, false);
    W("Skills Development (" + nv.SDWeight + " pts): training spend on black people as a % of leviable payroll, learnerships/internships/apprenticeships, learners absorbed. Priority sub-minimum applies.", 10, false);
    W("Enterprise & Supplier Development (" + nv.ESDWeight + " pts): procurement from B-BBEE compliant and black-owned suppliers, supplier development and enterprise development spend. Priority sub-minimum applies.", 10, false);
    W("Socio-Economic Development (" + nv.SEDWeight + " pts): contributions benefiting black people, as a % of net profit after tax.", 10, false);
    W("Levels (Generic): 100+ = Level 1, 95-99.99 = 2, 90-94.99 = 3, 80-89.99 = 4, 75-79.99 = 5, 70-74.99 = 6, 55-69.99 = 7, 40-54.99 = 8, below 40 = Non-compliant. Missing a priority sub-minimum drops the level by one.", 10, false);
    gap();
    W("GOOD QUESTIONS TO ASK", 12, true);
    W("'Explain my score'  -  'Which element is furthest from its maximum?'  -  'What is the cheapest way to the next level?'  -  'Which priority sub-minimums am I missing?'  -  'Explain Skills Development in plain words'", 10, false);
    gap();
    W("RULES", 12, true);
    W("Read-only. Cite the values you were given. Do not invent numbers. If the answer is not in the values provided, say so. Do not request formulas, documents or the full workbook. Mantis does not see your question in Mantis AI mode beyond a category and a count.", 10, false);
    gap();
    W("Deployed by " + ADDIN_NAME + " " + ADDIN_VERSION + " on " + longDate(new Date()), 9, false);

    ws.getRange().clear(Excel.ClearApplyTo.contents);
    ws.getRange().format.fill.clear();
    await ctx.sync();

    var first = 2, last = first + L.length - 1;
    var block = ws.getRange("B" + first + ":B" + last);
    block.values = L.map(function (x) { return [x ? x.t : ""]; });
    block.format.font.size = 10;
    block.format.font.bold = false;
    block.format.wrapText = true;
    block.format.verticalAlignment = Excel.VerticalAlignment.top;
    L.forEach(function (x, i) {
      if (!x) return;
      var c = ws.getRange("B" + (first + i));
      if (x.sz !== 10) c.format.font.size = x.sz;
      if (x.b) { c.format.font.bold = true; c.format.fill.color = LIGHT_FILL; }
    });
    ws.activate();
    ws.getRange("A1").select();
    await ctx.sync();
  }

  async function buildAISheet(ctx) {
    var ws = await addSheetAfter(ctx, AI_SHEET, GUIDE_SHEET);
    ws.getRange("A:A").format.columnWidth = charsToPoints(3);
    ws.getRange("B:B").format.columnWidth = charsToPoints(100);
    ws.getRange("C:C").format.columnWidth = charsToPoints(24);

    function put(addr, v) { var r = ws.getRange(addr); r.values = [[v]]; return r; }

    var t = put("B2", "Ask AI about this calculator"); t.format.font.size = 16; t.format.font.bold = true;
    var s = put("B3", "Only your question and the scorecard values are sent. Never formulas, never your documents. (POPIA)");
    s.format.font.size = 9; s.format.font.italic = true;

    put("B5", "Your question:").format.font.bold = true;
    var q = put("B6", "");
    q.format.fill.color = INPUT_FILL; q.format.wrapText = true; q.format.rowHeight = 60;
    q.format.verticalAlignment = Excel.VerticalAlignment.top; q.format.protection.locked = false;

    put("B8", "Answer:").format.font.bold = true;
    var a = put("B9", ANSWER_PLACEHOLDER);
    a.format.fill.color = LIGHT_FILL; a.format.wrapText = true; a.format.rowHeight = 300;
    a.format.verticalAlignment = Excel.VerticalAlignment.top; a.format.font.color = "#3C3C3C";

    var hint = put("C5", "Buttons: Home tab > Mantis Updater > AI");
    hint.format.font.size = 9; hint.format.font.italic = true; hint.format.wrapText = true;

    put("B11", "Settings").format.font.bold = true;
    put("B12", "Provider");
    var prov = put("C12", "Mantis AI");
    prov.format.fill.color = INPUT_FILL;
    prov.dataValidation.clear();
    prov.dataValidation.rule = { list: { inCellDropDown: true, source: "Mantis AI,My own key" } };
    prov.dataValidation.errorAlert = { showAlert: true, style: Excel.DataValidationAlertStyle.stop, title: ADDIN_NAME, message: "Choose Mantis AI or My own key." };
    put("B13", "My own API key (Anthropic sk-ant-... or OpenAI sk-...). Stored in this file only, never sent to Mantis.");
    var key = ws.getRange("C13");
    key.format.fill.color = INPUT_FILL;
    key.numberFormat = [[";;;"]];
    put("B14", "Status");
    put("C14", "ready");
    await ctx.sync();

    await setName(ctx, "AI_Question", ws.getRange("B6"));
    await setName(ctx, "AI_Answer", ws.getRange("B9"));
    await setName(ctx, "AI_Provider", ws.getRange("C12"));
    await setName(ctx, "AI_Key", ws.getRange("C13"));
    await setName(ctx, "AI_Status", ws.getRange("C14"));

    ws.activate();
    ws.getRange("B6").select();
    await ctx.sync();
  }

  // =====================================================================================
  // modMain - Apply updates, versions log, fetch patches
  // =====================================================================================
  async function readCalculator() {
    return Excel.run(function (ctx) { return calcInfo(ctx); });
  }

  async function applyUpdates() {
    return Excel.run(async function (ctx) {
      var info = await calcInfo(ctx);
      if (!info.isCalc) throw new Error("This workbook is not a Mantis calculator. Open the calculator first, then click Apply updates.");
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.8")) {
        ctx.workbook.load("readOnly");
        await ctx.sync();
        if (ctx.workbook.readOnly) throw new Error("The calculator is read-only. Save a copy you can write to, then try again.");
      }

      logLine("==== Apply updates: " + info.workbookName + " (" + info.product + " " + info.version + ")");
      var stage = "unprotect", pwd = info.password;
      try {
        var bp = ctx.workbook.protection; bp.load("protected"); await ctx.sync();
        var wasProt = bp.protected;
        var structOpen = await unprotectBookSafe(ctx, pwd);
        if (!structOpen) logLine("could not unprotect workbook structure - sheet-level patches will still be tried, new sheets will not");

        stage = "patches";
        var res = await applyPatchTable(ctx, info, pwd);

        stage = "AI sheets";
        var aiMsg = structOpen ? await ensureAISheets(ctx, info) : "AI sheets skipped (structure locked); ";

        stage = "versions log";
        await writeVersionsLog(ctx, info, pwd, ADDIN_NAME + " " + ADDIN_VERSION + ": " + res.applied + " patch(es) applied, " +
          res.skipped + " skipped. " + aiMsg + res.notes);

        stage = "save";
        if (wasProt && structOpen) await protectBookSafe(ctx, pwd);
        var saved = await saveWorkbook(ctx);

        logLine("done: applied=" + res.applied + " skipped=" + res.skipped);
        return { applied: res.applied, skipped: res.skipped, ai: aiMsg, saved: saved, product: info.product, version: info.version };
      } catch (e) {
        var d = "at stage '" + stage + "': " + (e && e.message ? e.message : e);
        logLine("ERROR " + d);
        throw new Error("Update stopped " + d);
      }
    });
  }

  // Appends a line to the calculator's own Versions sheet (Page | Comment | Version | Date)
  async function writeVersionsLog(ctx, info, pwd, comment) {
    try {
      var ws = await sheetOrNull(ctx, "Versions");
      if (!ws) return;
      ws.protection.load("protected"); await ctx.sync();
      var wasProt = ws.protection.protected;
      await unprotectSheetSafe(ctx, ws, pwd);
      var used = ws.getRange("C:C").getUsedRangeOrNullObject(true);
      used.load("rowIndex,rowCount");
      await ctx.sync();
      var r = used.isNullObject ? 1 : used.rowIndex + used.rowCount + 1;
      if (r < 7) r = 7;
      var row = ws.getRange("B" + r + ":E" + r);
      row.values = [[ADDIN_NAME, String(comment).slice(0, 32000), info.version, excelDateSerial(new Date())]];
      ws.getRange("E" + r).numberFormat = [["yyyy/mm/dd"]];
      await ctx.sync();
      if (wasProt) await protectSheetSafe(ctx, ws, pwd);
    } catch (e) {
      logLine("   WARNING: could not write the Versions log (" + e.message + ")");
    }
  }

  // Fetch the patch table CSV from Mantis (Settings: PatchURL). The server must allow CORS.
  async function fetchPatches() {
    var url = cfgValue("PatchURL");
    if (!url) throw new Error("No Patch URL in Settings. Using the local patch table.");
    var resp = await fetch(url, { cache: "no-store" });
    if (!resp.ok) throw new Error("HTTP " + resp.status);
    var rows = parsePatchCsv(await resp.text());
    setPatches(rows);
    logLine("fetched " + rows.length + " patch rows from " + url);
    return rows.length;
  }

  function loadPatchCsv(text, source) {
    var rows = parsePatchCsv(text);
    setPatches(rows);
    logLine("loaded " + rows.length + " patch rows from " + (source || "CSV"));
    return rows.length;
  }

  // =====================================================================================
  // modAI - Ask AI / Explain / Next level / Clear / Export snapshot
  // =====================================================================================
  var PRESETS = {
    explain: "Explain my current B-BBEE score in plain language: total points, level, and which elements are strong or weak. Keep it under 200 words.",
    next: "Using only the values provided, which indicators are furthest below their maximum points, and what are the three most practical moves to reach the next B-BBEE level? Do not invent numbers."
  };

  async function targetCalc(ctx) {
    var info = await calcInfo(ctx);
    if (!info.isCalc) throw new Error("This workbook is not a Mantis calculator.");
    if (!info.hasAISheet) throw new Error("This calculator has no AI sheet yet. Run 'Apply updates' first.");
    return info;
  }

  async function guideText(ctx) {
    var ws = await sheetOrNull(ctx, GUIDE_SHEET);
    if (!ws) return "";
    var r = ws.getRange("B2:B60"); r.load("values"); await ctx.sync();
    return r.values.map(function (v) { return textOf(v[0]); }).filter(function (s) { return s.length > 0; }).join("\n") + "\n";
  }

  async function scorecardRows(ctx) {
    var ws = await sheetOrNull(ctx, "Scorecard");
    if (!ws) return [];
    var r = ws.getRange("B10:I" + MAX_SCORECARD_ROWS); r.load("values,valueTypes"); await ctx.sync();
    return r.values.map(function (row, i) {
      return row.map(function (v, j) { return r.valueTypes[i][j] === "Error" ? "" : v; });
    });
  }

  async function scorecardValues(ctx) {
    var rows = await scorecardRows(ctx);
    var s = "";
    rows.forEach(function (row, i) {
      var line = row.map(round2).join(" | ");
      if (line.replace(/[| ]/g, "").length > 0) s += "row " + (10 + i) + ": " + line + "\n";
    });
    async function N(n) { return nameValue(ctx, n, true); }
    s += "Totalscore=" + (await N("Totalscore")) + "; BEELevel=" + (await N("BEELevel")) + "; BEELevel_After=" + (await N("BEELevel_After")) + "\n";
    s += "OwnScore=" + (await N("OwnScore")) + "/" + (await N("OWeight")) + "; MCScore=" + (await N("MCScore")) + "/" + (await N("MCWeight")) +
      "; SDScore=" + (await N("SDScore")) + "/" + (await N("SDWeight")) + "; ESDScore=" + (await N("ESDScore")) + "/" + (await N("ESDWeight")) +
      "; SEDScore=" + (await N("SEDScore")) + "/" + (await N("SEDWeight"));
    return s;
  }

  async function clientRef(ctx) {
    var v = await nameValue(ctx, "UniqueIdentifyingNumber");
    return v || (await nameValue(ctx, "Company_Name"));
  }

  // kind: "ask" (uses `question`, or the AI_Question cell if blank), "explain", "next"
  async function runQuestion(kind, question) {
    return Excel.run(async function (ctx) {
      var info = await targetCalc(ctx);
      var q = PRESETS[kind] || trim(question);
      if (q) await setNameValue(ctx, "AI_Question", q);
      else q = trim(await nameValue(ctx, "AI_Question"));
      if (!q) throw new Error("Type a question first.");

      var mode = (await nameValue(ctx, "AI_Provider")) || "Mantis AI";
      await setNameValue(ctx, "AI_Status", "thinking...");

      var sys = (await guideText(ctx)) + "\nSCORECARD VALUES (indicator | detail... | points available | points scored):\n" + (await scorecardValues(ctx));
      var key = await nameValue(ctx, "AI_Key");
      var ans = await callProvider(mode, key, sys, q, {
        product: info.product, version: info.version, client: await clientRef(ctx)
      });

      await setNameValue(ctx, "AI_Answer", ans.slice(0, 32000));
      var now = new Date();
      await setNameValue(ctx, "AI_Status", "done " + pad(now.getHours()) + ":" + pad(now.getMinutes()));
      logLine("AI usage: " + info.product + " " + info.version + " mode=" + mode + " category=" + questionCategory(q));
      return { question: q, answer: ans, mode: mode };
    });
  }

  async function postJson(url, headers, body) {
    var ctrl = typeof AbortController !== "undefined" ? new AbortController() : null;
    var timer = ctrl ? setTimeout(function () { ctrl.abort(); }, 60000) : null;
    try {
      var resp = await fetch(url, { method: "POST", headers: headers, body: JSON.stringify(body), signal: ctrl ? ctrl.signal : undefined });
      return { status: resp.status, ok: resp.ok, text: await resp.text() };
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  // Never throws - returns the answer or a friendly error, like the VBA version
  async function callProvider(mode, key, sys, q, meta) {
    try {
      var r;
      if (mode === "My own key") {
        key = trim(key);
        if (!key) return "Enter your API key in the settings cell first (AI sheet, C13).";
        if (key.indexOf("sk-ant-") === 0) {
          r = await postJson("https://api.anthropic.com/v1/messages", {
            "content-type": "application/json",
            "x-api-key": key,
            "anthropic-version": "2023-06-01",
            "anthropic-dangerous-direct-browser-access": "true"
          }, { model: cfgValue("AnthropicModel") || DEFAULT_CONFIG.AnthropicModel, max_tokens: 800, system: sys, messages: [{ role: "user", content: q }] });
        } else {
          r = await postJson("https://api.openai.com/v1/chat/completions", {
            "content-type": "application/json",
            "Authorization": "Bearer " + key
          }, { model: cfgValue("OpenAIModel") || DEFAULT_CONFIG.OpenAIModel, max_tokens: 800,
            messages: [{ role: "system", content: sys }, { role: "user", content: q }] });
        }
      } else {
        var url = cfgValue("MantisAIEndpoint");
        if (!url) return "Mantis AI is not switched on yet (no endpoint configured). Choose 'My own key' or use Export AI snapshot.";
        // client ref = product + registration key, so the proxy can check the AI Tier entitlement
        r = await postJson(url, { "content-type": "application/json" },
          { product: meta.product, version: meta.version, client: meta.client, system: sys, question: q });
      }
      if (!r.ok) return "AI unavailable (HTTP " + r.status + "). Your calculator is unaffected.\n" + r.text.slice(0, 300);
      var ans = extractAnswer(r.text);
      return ans || "No answer returned.\n" + r.text.slice(0, 300);
    } catch (e) {
      return "AI unavailable (" + (e && e.name === "AbortError" ? "timed out" : (e && e.message) || e) + "). Your calculator is unaffected.";
    }
  }

  async function clearAI() {
    return Excel.run(async function (ctx) {
      await targetCalc(ctx);
      await setNameValue(ctx, "AI_Question", "");
      await setNameValue(ctx, "AI_Answer", ANSWER_PLACEHOLDER);
      await setNameValue(ctx, "AI_Status", "ready");
    });
  }

  // Values-only snapshot opened as a NEW workbook (the user saves it and uploads it to Claude/ChatGPT)
  async function exportSnapshot() {
    if (typeof XLSX === "undefined") throw new Error("The spreadsheet library did not load - check your internet connection.");
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.8")) throw new Error("This version of Excel cannot open a new workbook from an add-in.");
    var data = await Excel.run(async function (ctx) {
      var info = await targetCalc(ctx);
      return { info: info, guide: await guideText(ctx), rows: await scorecardRows(ctx) };
    });
    var aoa = [
      ["Mantis " + data.info.product + " " + data.info.version + " - values-only snapshot for AI (no formulas, no client documents)"],
      ["Instructions for the AI: " + data.guide],
      [],
      ["Scorecard values:"]
    ].concat(data.rows);
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 60 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 }];
    var nb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(nb, ws, "AI Snapshot");
    var b64 = XLSX.write(nb, { bookType: "xlsx", type: "base64" });
    await Excel.createWorkbook(b64);
    var suggested = "AI Snapshot " + data.info.product + " " + fileStamp(new Date()) + ".xlsx";
    logLine("AI snapshot opened for " + data.info.product + " " + data.info.version);
    return suggested;
  }

  return {
    ADDIN_NAME: ADDIN_NAME, ADDIN_VERSION: ADDIN_VERSION, PATCH_COLUMNS: PATCH_COLUMNS,
    // storage
    getConfig: getConfig, setConfig: setConfig, getPatches: getPatches, setPatches: setPatches,
    getLog: getLog, clearLog: clearLog, onLog: onLog, logLine: logLine,
    // pure helpers
    compareVersion: compareVersion, normHex: normHex, replaceText: replaceText, splitCsvLine: splitCsvLine,
    parsePatchCsv: parsePatchCsv, patchesToCsv: patchesToCsv, patchApplies: patchApplies, questionCategory: questionCategory,
    extractAnswer: extractAnswer, cellAddress: cellAddress, round2: round2,
    // Excel actions
    readCalculator: readCalculator, applyUpdates: applyUpdates, fetchPatches: fetchPatches, loadPatchCsv: loadPatchCsv,
    runQuestion: runQuestion, clearAI: clearAI, exportSnapshot: exportSnapshot
  };
})();

// let Node load this file for the tests
if (typeof module !== "undefined") module.exports = MU;
