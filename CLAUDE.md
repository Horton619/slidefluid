# SlideFluid

> **Project-scoped CLAUDE.** Loads when working in `~/slidefluid`.
> Cross-project lessons (signing, notarization, auto-update, design language) live in `~/.claude/CLAUDE.md` (the "global CLAUDE").

SlideFluid is a Mac/Windows Electron desktop app that converts files into PPTX presentations. Three input pipelines: **PDF → PPTX** (rasterized with pillarbox fill modes), **DOCX/TXT → PPTX** (text content with auto-fit font sizing), and **PPTX → notes pipeline** (teleprompter `.txt`, split deck, or paired audience+companion deck). Single Python file (~2000 lines) holds all backend logic; renderer is no-bundler vanilla JS. Internal distribution for AVSC / Visual Entropy Productions, signed + notarized with the VEP-wide Developer ID (team `L5KZ5KGKXC`).

---

## ⚠ READ FIRST — load the right topic doc before writing code

| If your task touches… | Read first |
|---|---|
| `_pptx_clone_slide`, deep-copy, chart/SmartArt/OLE rels, chunking, threshold/soft-buffer, `chunks[0] × (builds + 1)`, animations, companion sync, `_atomic_save_pptx`, anything in the PPTX→notes pipeline | [`docs/PPTX_PIPELINE.md`](docs/PPTX_PIPELINE.md) |
| `_parse_txt`, `_parse_docx`, `_strip_rtf`, font fit constants (`_CHAR_W_RATIO`/`_LINE_H_RATIO`), `_add_text_slide`, `docx_analyze`, the overflow modal, `--split-choices`, `--text-align`, `--slide-theme` | [`docs/DOCX_TXT.md`](docs/DOCX_TXT.md) |
| `convert_pdf`, `fill_black/fill_color_match/fill_smear`, `_bar_dims`, DPI defaults, vendored Poppler, `resolvePopplerPath`, the renderer's fill-mode picker | [`docs/PDF_FILL.md`](docs/PDF_FILL.md) |
| `ConversionJob`, `_handleMessage`, `_buildArgs`, NDJSON message types, `beginConversion` grouping, `preload.js` surface, `_emit`, adding/renaming any CLI flag or IPC channel | [`docs/IPC.md`](docs/IPC.md) |
| Code signing, notarization, electron-updater, GitHub releases, the `release.yml` workflow, `entitlements.mac.plist` | `~/.claude/CLAUDE.md` (global) — already comprehensive, don't duplicate here |

If your task touches multiple areas, load multiple docs. They're sized small (50–150 lines each) on purpose; loading two is cheap.

---

## Universal rules (apply everywhere — no doc-load required)

- **Don't commit until asked.** When asked, draft a focused conventional-commit message. Never `git push` without explicit instruction.
- **Don't refactor / extract / rename beyond the requested change.** Three similar lines is better than a premature abstraction.
- **Don't add error handling around internal code.** Validate at boundaries only (IPC, file load, GitHub API). Trust internally.
- **Don't add "what" comments.** Only "why" — hidden constraints, workarounds, surprises.
- **Always use `venv/bin/python3` in dev**, never system `python3`. Install new Python deps with `venv/bin/pip`.
- **Trust git log + file tree over any summary,** including this one. The renderer (`app.js`) reshapes fastest; re-read after ~5 turns of edits there.
- **Verify python-pptx / python-docx call signatures before assuming them.** Both libraries have bit us — see PPTX_PIPELINE and DOCX_TXT docs for the specific traps.
- **Atomic write for any conversion output.** Temp suffix + `os.replace`. Never write directly to the final path.

---

## Stack

- **Frontend:** Electron + vanilla HTML/CSS/JS renderer. No bundler, no TypeScript.
- **Backend:** Python 3 subprocess. Deps: `numpy`, `Pillow`, `pdf2image`, `python-pptx` (1.0.x), `python-docx`, `lxml`.
- **Vendored Poppler** in `vendor/poppler/{mac,win,linux}/`. Dev falls back to system Poppler; packaged builds always use vendored. See `docs/PDF_FILL.md`.
- **Packaging:** PyInstaller for the backend binary + electron-builder for the app shell. CI workflow `.github/workflows/release.yml` builds mac-arm64 (signed + notarized) and win-x64 (unsigned) on `v*.*.*` tag push.
- **Auto-updater:** `electron-updater` reads GitHub Releases (`latest-mac.yml` / `latest.yml`). Both `dmg` AND `zip` targets required on Mac for in-place updates — see global CLAUDE for the recipe.

---

## Current state

Latest release v3.1.7 (May 2026), signed + notarized end-to-end. All three pipelines (PDF, DOCX/TXT, PPTX→notes) are shipped and live. Most recent feature work was the PPTX→notes pipeline (`feat: PPTX → notes pipeline`, commit `2164c37`) — teleprompter / split deck / companion modes. Stable; no known broken paths.

---

## Project structure

```
backend/
  slidefluid_convert.py         The whole backend. PDF, DOCX/TXT, PPTX-notes pipelines.   → docs/PDF_FILL.md
                                                                                          → docs/DOCX_TXT.md
                                                                                          → docs/PPTX_PIPELINE.md
                                                                                          → docs/IPC.md
  slidefluid_convert.spec       PyInstaller spec — produces single-file binary
src/
  main.js                       Electron main: BrowserWindow, ConversionJob, IPC,         → docs/IPC.md
                                settings, logger, auto-updater wiring, Poppler resolve    → docs/PDF_FILL.md
  preload.js                    contextBridge surface — the renderer's IPC API            → docs/IPC.md
  renderer/
    index.html                  Layout, modals, banners, the update toast
    styles.css                  Pro / Fun skin via CSS custom properties
    app.js                      Queue, sidebar dispatcher, overflow modal,                → docs/IPC.md
                                conversion flow, sidebar per-fileType controls            → docs/DOCX_TXT.md
                                                                                          → docs/PPTX_PIPELINE.md
vendor/poppler/{mac,win,linux}/ Per-platform pdfinfo + pdftoppm + libs                    → docs/PDF_FILL.md
build/entitlements.mac.plist    Hardened-runtime exceptions for the PyInstaller backend   → global CLAUDE
assets/icons/                   App icon (icns + ico)
.github/workflows/release.yml   Tag-driven signed/notarized CI                            → global CLAUDE
```

---

## Backend ↔ Renderer protocol (condensed)

Backend is a Python subprocess. One JSON object per stdout line. Recognized types: `start`, `progress`, `done`, `warn`, `error`, `batch_done`, plus one-shot probes (`docx_info`, `pptx_info`, `docx_analyze`, `preflight_result`). Renderer groups queued items by their per-type config (PDFs by `fillMode`; DOCX/TXT by `slideTheme|textAlign`; PPTX by `mode|threshold|headers|slidenums|theme`), one spawn per group, awaits `batch_done` between groups.

Full schema, per-type payload shapes, group keys, CLI flag inventory, and "do NOT" rules: **[`docs/IPC.md`](docs/IPC.md)**.

---

## Environment & dev

- **`venv/bin/python3`** in dev (never system `python3`). Deps in `venv/`. New deps: `venv/bin/pip install <pkg>`.
- **Vendored Poppler** for dev too (`vendor/poppler/<platform>/`); falls back to system PATH if absent.
- **Build PyInstaller binary** for local-packaged testing: `cd backend && pyinstaller --clean slidefluid_convert.spec`.
- **`npm start`** runs Electron in dev (`autoUpdater` is a no-op here — only fires in packaged builds).
- **CI secrets** for signing/notarization: `CSC_LINK`, `CSC_KEY_PASSWORD`, `APPLE_ID`, `APPLE_APP_SPECIFIC_PASSWORD`, `APPLE_TEAM_ID`. Set on `Horton619/slidefluid` repo. See global CLAUDE for the recipe.

---

## Universal quirks (cross-cutting; not big enough to be a topic doc)

- **Logger writes to `~/Library/Application Support/slidefluid/slidefluid.log`** (Mac) / equivalent on Windows. Rotates at 2MB. `ConversionJob._handleMessage` mirrors `warn`/`error` events into the log so post-mortem diagnosis is sufficient without re-running with stdout capture.
- **Settings store at `userData/slidefluid-settings.json`.** Schema is loose-typed; new keys safe to add. Defaults applied at load time, so a missing key in the on-disk file uses the in-code default.
- **Two skins: Professional (`#070910` + `#3DFFCC` teal) and Fun (trans-flag colors).** Skin = CSS custom-property override. The Fun skin also has rotating progress strings + confetti on `batch_done`.
- **Drop-a-file auto-selects it** so accidental re-conversions of completed items don't happen.
- **Per-mode output filenames are deliberately distinct** (`_notes.txt`, `_split.pptx`, `_audience.pptx`, `_companion.pptx`). They're user-facing nomenclature. Don't template them.
- **Done items get a clickable ⓘ** that opens a Conversion Details modal showing the `edits[]` payload. `Done ⚠ N` for items with warnings. Same modal either way.
- **DPI default is 144** (was 72; updated after real complaints about pixelated text-heavy slides). See `docs/PDF_FILL.md`.

---

## Branding / design

- **Pro skin tagline:** `"Transition to greatness"` (static).
- **Fun skin tagline:** pool of 7 in `FUN_TAGLINES`, one picked per app launch — includes "How does your PPT identify?" Format-reassignment tone throughout.
- **Footer:** `position: fixed; bottom: 0; z-index: 50` — "© 2026 Visual Entropy Productions" + `veproductions.net` link. No AVSC reference; SlideFluid is strictly VEP.
- **Per-app accent:** SlideFluid is `#3DFFCC`. Status colors follow global VEP semantics — see global CLAUDE.

---

## Release flow

1. Bump `package.json` `"version"`
2. Commit (conventional-commit message)
3. `git push origin main`
4. `git tag vX.Y.Z && git push origin vX.Y.Z`
5. CI builds + signs + notarizes + publishes to GitHub Releases on the tag

Users on the previous version auto-update on next cold launch (within GitHub's atom-feed cache window, ~5 min). See global CLAUDE for the full auto-update recipe.

---

## Working agreement

### Ask before doing when
- Change spans 3+ files. One-paragraph plan + file list first.
- About to change a documented decision in any topic doc (the "do NOT" lists are load-bearing).
- Touching `_pptx_clone_slide`, the chunking math, the font-fit constants, the IPC schema, or the group-key rules.

### Just do it when
- One-file diff, mechanical change, no design decisions.
- Bug fix where the symptom + traceback name the cause.
- Cosmetic / copy tweak in already-shipped UI.

### Always re-read source when
- It's been more than ~5 turns since you last read the file. The renderer (`app.js`) changes shape fastest.
- About to assume a python-pptx or python-docx signature. Both libs have surprises — see the topic docs.
- A summary handoff tells you what's there. Trust git log and the file tree, not the summary.

### Flag uncertainty
Saying *"I'm not sure if this also affects X — let me check"* is welcomed. Confident tone on a guess is the failure mode.

---

## Reference paths

- **PLAN.md** — original v3.0 build plan. Historical only; not actively maintained. The features list in this file's "Current state" supersedes it.
- **Crew letter samples (DOCX_TXT calibration data):**
  `/Users/horton/Desktop/TSCRA 26/TSCRA_2026_Crew_Letter.txt`
  `/Users/horton/Library/CloudStorage/Dropbox/Misc Show prod info/PAT 25/PAT 25 Crew Letter.rtf`
  `/Users/horton/Library/CloudStorage/Dropbox/Misc Show prod info/NCARB 25/NCARB ABM 25 Crew Letter.rtf`
  `/Users/horton/Library/CloudStorage/Dropbox/Misc Show prod info/NCARB LAS 25/LAS 25 Crew Letter.rtf`
