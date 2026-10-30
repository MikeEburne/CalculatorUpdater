// Run with:  node tests.js   (checks the helpers ported from modUtil / modPatch / modAI)
var assert = require("assert");
var MU = require("./updater.js");
var n = 0;
function t(name, fn) { fn(); n++; console.log("ok  " + name); }

t("compareVersion", function () {
  assert.strictEqual(MU.compareVersion("V15.10.0", "V15.9.9"), 1);
  assert.strictEqual(MU.compareVersion("v15.0.0", "V15.0"), 0);
  assert.strictEqual(MU.compareVersion("V4.11..0", "V4.11.1"), -1);
});

t("normHex", function () {
  assert.strictEqual(MU.normHex("ff0000"), "#FF0000");
  assert.strictEqual(MU.normHex("#00ff00"), "#00FF00");
  assert.strictEqual(MU.normHex("0x0000ff"), "#0000FF");
  assert.strictEqual(MU.normHex("red"), null);
  assert.strictEqual(MU.normHex(""), null);
});

t("replaceText - case-insensitive, all matches, idempotent", function () {
  assert.strictEqual(MU.replaceText("Black Women", "", "New"), "New");
  assert.strictEqual(MU.replaceText("Skills dev, skills", "skills", "SD"), "SD dev, SD");
  assert.strictEqual(MU.replaceText("Already SD", "skills", "SD"), "Already SD");
  assert.strictEqual(MU.replaceText("cost $1", "$1", "$2"), "cost $2");
});

t("splitCsvLine / parsePatchCsv / patchesToCsv round trip", function () {
  assert.deepStrictEqual(MU.splitCsvLine('a,"b,c","d ""q"""'), ["a", "b,c", 'd "q"']);
  var csv = "PatchID,Product,MinVer,MaxVer,Type,Sheet,Range,Find,Replace,FillHex,FontHex,Expect,Note\r\n" +
    'P-1,DTiAG,V15.0.0,,Rename,Scorecard,"B86,B100",Old,New,,,,"fix, label"\r\n\r\n';
  var rows = MU.parsePatchCsv(csv);
  assert.strictEqual(rows.length, 1);
  assert.strictEqual(rows[0].Range, "B86,B100");
  assert.strictEqual(rows[0].Note, "fix, label");
  assert.deepStrictEqual(MU.parsePatchCsv(MU.patchesToCsv(rows)), rows);
});

t("patchApplies", function () {
  var p = { Product: "DTiAG", MinVer: "V15.0.0", MaxVer: "V15.13.0" };
  assert.ok(MU.patchApplies(p, "dtiag", "V15.13.0"));
  assert.ok(!MU.patchApplies(p, "DTiAG", "V15.14.0"));
  assert.ok(!MU.patchApplies(p, "ConsAG", "V15.1.0"));
  assert.ok(MU.patchApplies({ Product: "*" }, "anything", "V1"));
});

t("questionCategory", function () {
  assert.strictEqual(MU.questionCategory("How do I reach the next level?"), "next-level");
  assert.strictEqual(MU.questionCategory("Explain my score"), "explain");
  assert.strictEqual(MU.questionCategory("ownership?"), "element");
  assert.strictEqual(MU.questionCategory("hello"), "other");
});

t("extractAnswer - Anthropic, OpenAI, proxy, fallback", function () {
  assert.strictEqual(MU.extractAnswer(JSON.stringify({ content: [{ type: "text", text: "Hi\nthere" }] })), "Hi\nthere");
  assert.strictEqual(MU.extractAnswer(JSON.stringify({ choices: [{ message: { content: "Yo" } }] })), "Yo");
  assert.strictEqual(MU.extractAnswer(JSON.stringify({ answer: "A" })), "A");
  assert.strictEqual(MU.extractAnswer('junk "text":"x\\"y" more'), 'x"y');
  assert.strictEqual(MU.extractAnswer("nothing"), "");
});

t("cellAddress / round2", function () {
  assert.strictEqual(MU.cellAddress(85, 1), "B86");
  assert.strictEqual(MU.cellAddress(0, 26), "AA1");
  assert.strictEqual(MU.cellAddress(0, 255), "IV1");
  assert.strictEqual(MU.round2(12.3456), "12.35");
  assert.strictEqual(MU.round2("7"), "7");
  assert.strictEqual(MU.round2("Level 4"), "Level 4");
});

console.log("\n" + n + " groups passed");
