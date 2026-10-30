# Mantis Updater (Office.js add-in)

This is the JavaScript version of `Mantis Updater V1.0.0.xlam`. It's built the same way as Quick Sheet: plain files on GitHub Pages, with no Node.js or build step. Follow `MS Store\quicksheet\quicksheet\PUBLISH-GUIDE.md`, using the repository name **`mantis-updater`**. The manifest already points at `https://mikeeburne.github.io/mantis-updater/`.

| File | What it is |
|---|---|
| `manifest.xml` | Add-in manifest. It adds a **Mantis Updater** button to the Home tab. |
| `taskpane.html` | The pane, with Update, AI, Settings and Log tabs. It replaces the "Mantis Updates" menu. |
| `updater.js` | All the logic. It's a port of modMain, modPatch, modAIGuide, modAI and modUtil. |
| `commands.html`, `privacy.html`, `support.html`, `assets/` | Required by the manifest and the store. |
| `tests.js` | `node tests.js` checks the ported helpers. |

## What maps to what

| VBA (.xlam) | Office.js add-in |
|---|---|
| Menu: Apply updates / Fetch patches / Show patch table / Show log / About | Pane: **Update** tab, **Log** tab, **Settings > About** |
| `Patches` sheet | Patch table in the pane. Load a CSV or use **Fetch from Mantis**. Same 13 columns. |
| `Config` sheet (`PatchURL`, `MantisAIEndpoint`) | **Settings** tab, plus the AI model names |
| `Log` sheet | **Log** tab (with Copy log) |
| Format / Value / Formula / Rename patches, Expect guard, formula/error-cell protection | The same rules, unchanged |
| Applied-patch register (`MantisPatch_<id>` custom document properties) | The same properties, so calculators already patched by the .xlam are recognised |
| Unprotect with `Dashboard!IV1`, re-protect afterwards | The same |
| AI Guide + AI sheets, names `AI_Question`, `AI_Answer`, `AI_Provider`, `AI_Key`, `AI_Status` | The same sheets and names |
| Buttons on the AI sheet (Ask AI, Explain, Next level, Export, Clear) | Buttons on the pane's **AI** tab. The sheet shows a hint pointing to them. |
| Versions log line | The same |
| Export AI snapshot (saved .xlsx next to the calculator) | Opens the snapshot as a new workbook, and the user saves it |
| AI on Mac not supported | Works on Windows, Mac and Excel on the web |

## What the add-in cannot do (Office.js limits)

1. **VBModule / VBRemove rows are skipped** and logged with a reason. Office add-ins have no access to a workbook's VBA project. Keep the .xlam for code patches, or ship those changes in a new calculator release.
2. **No buttons on the sheet.** Add-ins can't create form buttons that run code. Calculators that the .xlam already patched still have buttons pointing at the .xlam macros. They stop working once the .xlam is gone, so a small Format/Value patch or a new release should remove or relabel them.
3. **No backup copy** before patching. Add-ins can't save a copy to disk. This was switched off in the prototype as well.
4. **One workbook at a time.** The add-in works on the workbook its pane is open in, like "active workbook" in the VBA.

## Things to set up before going live

- **CORS:** the Patch URL and the Mantis AI endpoint must send `Access-Control-Allow-Origin: https://mikeeburne.github.io` (or `*`), or the browser blocks the request. Anthropic and OpenAI already allow browser calls; the add-in sends Anthropic's `anthropic-dangerous-direct-browser-access` header.
- **Model names:** the defaults are taken from the VBA (`claude-sonnet-4-5`, `gpt-4o-mini`). Change them on the Settings tab if needed.
- **Excel version:** the add-in needs Excel API 1.7 or later (Excel 2019+, Microsoft 365 and the web). Automatic save needs API 1.11. On older versions the pane asks the user to press Ctrl+S.
