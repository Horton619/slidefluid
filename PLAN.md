# SlideFluid 3.0 — Build Plan

## What exists (do not rebuild)

| File | Status | Notes |
|------|--------|-------|
| `slidefluid_convert.py` | ✅ Complete | Python conversion engine — PDF→PPTX, 3 fill modes, IPC JSON output, preflight, batch |
| `main.js` | ✅ Complete | Electron main process — BrowserWindow, ConversionJob, settings, IPC handlers |
| `preload.js` | ✅ Complete | contextBridge — full `window.slidefluid` API exposed to renderer |
| `package.json` | ✅ Complete | Electron + electron-builder config for Mac (.dmg) and Win (.exe) |

---

## Phase 1 — Project structure & wiring

**Goal:** `npm start` opens a blank window without errors.

### Tasks

- [x] **1.1** Create directory skeleton
  ```
  src/            ← main.js and preload.js move here
  src/renderer/   ← index.html, styles.css, app.js go here
  backend/        ← slidefluid_convert.py moves here
  assets/icons/   ← placeholder icons
  vendor/poppler/ ← (leave empty for now; dev uses system Poppler)
  ```

- [x] **1.2** Move files to correct locations
  - `main.js` → `src/main.js`
  - `preload.js` → `src/preload.js`
  - `slidefluid_convert.py` → `backend/slidefluid_convert.py`
  - Verify `package.json` `"main": "src/main.js"` (already correct)

- [x] **1.3** Add `pdf:info` and `pdf:scan` IPC handlers to `src/main.js`
  - `pdf:info(filePath)` — calls `pdfinfo` (Poppler), returns `{pageCount, ar, widthPt, heightPt}`
  - `pdf:scan(paths[])` — expands any folder paths to a flat list of PDF paths
  - Both exposed in `src/preload.js` as `getPdfInfo(path)` and `scanPaths(paths[])`

- [x] **1.4** Create stub `src/renderer/index.html` (just enough to confirm window loads)

- [ ] **1.5** Run `npm install` — installs Electron and electron-builder

- [ ] **1.6** Run `npm start` — confirm window opens, DevTools loads, no console errors

---

## Phase 2 — Renderer: layout & drop zone

**Goal:** UI shell visible, files can be dropped and appear in queue.

### Spec refs
- Section 5.1: Window 800×560, min 640×440, header + left pane + right sidebar
- Section 5.2: Right sidebar layout
- Section 4.2: Batch modes (multi-file, folder)

### Tasks

- [x] **2.1** Write `src/renderer/index.html`
  - Header bar (logo placeholder, app name, gear icon)
  - Left pane: drop zone div, queue list div, progress section div, convert button
  - Right sidebar: preview canvas (320×180), AR badge, fill mode selector, output folder row, file details

- [x] **2.2** Write `src/renderer/styles.css` — Professional skin only
  - CSS custom properties for Professional palette (`#070910` bg, `#3DFFCC` accent, etc.)
  - Layout: header fixed height, body as flex row, left pane fixed width ~420px, sidebar fills rest
  - Drop zone: dashed accent border, centered label, hover/drag-over state
  - Queue item: filename, status badge (Waiting/Converting/Done/Error), page count
  - Progress bars: thin 3px accent color

- [x] **2.3** Write `src/renderer/app.js` — drop zone & queue only
  - State: `queue[]`, `selectedId`, `outputDir`, `settings`, `isConverting`
  - Drop zone: dragover, dragleave, drop events; calls `window.slidefluid.scanPaths()` for folders
  - File picker button: calls `window.slidefluid.openFilePicker()`
  - `addFiles(paths[])` — deduplicates, creates queue items, calls `pdf:info` per file
  - `renderQueue()` — builds queue item DOM
  - `selectItem(id)` — highlights item, updates sidebar
  - Preview canvas: draws grey gradient rect + colored bars based on AR + fill mode (no PDF.js yet)
  - AR badge: "16:9 — correct" (green) or "4:3 — pillarbox" (amber)
  - Fill mode selector: shown only for non-16:9 files; Black | Color Match | Smear Fill

- [ ] **2.4** Test: drop a 16:9 PDF → appears in queue with page count, AR badge green
- [ ] **2.5** Test: drop a 4:3 PDF → fill mode selector appears, preview shows bars
- [ ] **2.6** Test: drop a folder → all PDFs inside added individually to queue

---

## Phase 3 — Renderer: output folder & convert button

**Goal:** Can pick output folder; Convert button wires up to Python backend end-to-end.

### Spec refs
- Section 4.3: Output folder selection — modal prompt on first drop, "Change" button
- Section 4.5: Convert button 4 states
- Section 4.6: Cancel / abort

### Tasks

- [x] **3.1** Output folder row
  - On app load: check `settings.outputDir`; if null show "No output folder set" + "Choose Folder" button (amber)
  - If set: show path + "Change Folder" button
  - `window.slidefluid.openFolderPicker()` saves to settings automatically
  - First-drop modal: if no outputDir set when files are first dropped, show folder picker immediately

- [x] **3.2** Convert button state machine
  - **Idle** (no files or no outputDir): dimmed, label "Convert to PPTX"
  - **Ready** (files queued + folder set): active, accent fill, label "Convert to PPTX"
  - **Converting**: label "Cancel", click → `cancelConversion()`
  - **Complete**: label "Open Output Folder", click → `window.slidefluid.openFolder(outputDir)`
  - State resets to Ready when new file is dropped after completion

- [x] **3.3** Pre-conversion overwrite check
  - Compute output path for each file: `outputDir / (stem + suffix + '.pptx')`
  - Call `window.slidefluid.checkOverwrite(path)` for each
  - Handle: overwrite (keep), skip (remove from batch), cancel (abort all)
  - 'rename' behavior: walk `_1`, `_2`, … until no conflict
  - Note: dialogs appear per-file for 'ask' mode (consolidated dialog is a Phase 6 polish item)

- [x] **3.4** Start conversion
  - Call `window.slidefluid.startConversion({files, outputDir, dpi, fillMode, suffix})`
  - Subscribe to `onConversionMessage`, `onConversionStderr`, `onConversionExit`, `onConversionSpawnError`

- [x] **3.5** IPC message handling
  - `start`: mark queue item as Converting, show per-file progress bar
  - `progress`: update per-file progress bar, overall bar, and status text
  - `done`: mark queue item Done, store output path
  - `error`: mark queue item Error, show message
  - `batch_done`: show summary line, update convert button to Complete state, auto-open if setting on
  - `warn`: show inline warning on queue item

- [x] **3.6** Cancel
  - Call `window.slidefluid.cancelConversion()`
  - Delete partial PPTX: Python handles this internally
  - Reset in-flight item to Waiting; completed items stay Done
  - Convert button resets to Ready

- [x] **3.7** Test: convert a 16:9 PDF → PPTX created, progress updates, button → "Open Output Folder"
- [x] **3.8** Test: convert a 4:3 PDF with each fill mode → check PPTX output visually in PowerPoint
- [x] **3.9** Test: cancel mid-batch → partial file deleted, completed files remain
- [x] **3.10** Test: batch of 3 files → summary line correct

---

## Phase 4 — Settings modal

**Goal:** Gear icon opens modal; all settings tabs functional.

### Spec refs
- Section 6.1: Appearance tab
- Section 6.2: Output tab
- Section 6.3: Diagnostics tab

### Tasks

- [ ] **4.1** Settings modal shell
  - Modal overlay (backdrop + centered panel)
  - Three tabs: Appearance | Output | Diagnostics
  - Close on backdrop click or Escape key
  - Gear icon in header opens modal

- [ ] **4.2** Appearance tab
  - Skin toggle: Professional (FlowCast) / Fun (SlideFluid Classic)
  - Changing skin updates `data-skin` on `<html>` and persists via `setSetting('skin', ...)`

- [ ] **4.3** Output tab
  - Default output folder (path display + clear button)
  - Output DPI: 72 (default) / 144 radio — show file size warning for 144
  - Default pillarbox fill: Black / Color Match / Smear Fill
  - Auto-open output folder on complete: toggle
  - Overwrite behavior: Ask / Always Overwrite / Always Skip / Always Rename
  - Filename suffix: text input

- [ ] **4.4** Diagnostics tab
  - "Run Preflight Check" button → calls `window.slidefluid.runPreflight()`
  - Subscribes to `onPreflightResult`, renders pass/fail table:
    - Poppler binary
    - Python subprocess
    - python-pptx
    - Pillow / pdf2image
    - Output folder writable
    - Disk space
    - App version
  - "Copy to Clipboard" button → copies report as plain text

- [ ] **4.5** Test: all settings persist across app restart
- [ ] **4.6** Test: preflight shows all green on dev machine
- [ ] **4.7** Test: DPI change is respected in next conversion

---

## Phase 5 — Fun skin

**Goal:** Toggling to Fun skin changes all text, colors, and animations.

### Spec refs
- Section 5.3: Skin B — Internal/Fun
- Section 5.4: Feature comparison table

### Tasks

- [ ] **5.1** CSS — Fun skin palette
  - `[data-skin="fun"]` overrides CSS custom properties
  - Trans flag colors: `#55CDFC` (blue), `#F7A8B8` (pink), `#FFFFFF` (white) on `#0D0D14`
  - Gradient shimmer animation on drop zone label text

- [ ] **5.2** Skin-aware text strings in app.js
  - `skinText(professionalStr, funStr)` helper reads current skin from state
  - Drop zone label: "Drop PDFs here or click to browse." / rotating one-liners
  - Progress messages: "Converting page N of M" / rotating witty strings
  - Convert button (idle/ready): "Convert to PPTX" / "Set them free →"
  - Convert button (complete): "Open Output Folder" / "They're free. Open folder?"
  - Settings label: "Settings" / "Options & Stuff"

- [ ] **5.3** Fun rotating strings
  - Drop zone one-liners: array, cycles on each new file drop
  - Progress messages: array, cycles per page
  - Confetti burst on `batch_done`: CSS keyframe animation, brief and subtle

- [ ] **5.4** App icon differentiation (section 5.3 mentions different icon per skin)
  - In-app window title only — both skins share same OS-level app icon
  - Window title: "SlideFluid" (Pro) / "SlideFluid ✦" (Fun)

- [ ] **5.5** Test: toggle skin mid-session → all text and colors update instantly, no reload

---

## Phase 6 — Polish, resize & future scaffolding

**Goal:** All error paths handled; app feels finished; DOCX and live preview scaffolding in place.

### Spec refs
- Section 8: Error handling
- Section 4.5: Progress feedback details

### Tasks

- [ ] **6.1** Drop zone rejection — DOCX-aware
  - `.pdf` files accepted normally
  - `.docx` files show a friendly inline notice: "DOCX → PPTX coming in a future update" — never added to queue
  - All other file types show brief red flash on drop zone border then dismiss
  - Add `.docx` to `openFilePicker` accepted extensions (so Finder doesn't hide them)

- [ ] **6.2** Empty queue state
  - Drop zone visible and prominent when queue is empty
  - Fade queue area in when first file is added

- [ ] **6.3** File details panel (right sidebar)
  - Filename (truncated with tooltip if long)
  - Page count (from `pdf:info`)
  - Detected AR (from `pdf:info`)
  - Estimated output size (rough: pages × 200-500 KB at 72 DPI)

- [ ] **6.4** Overall progress bar
  - "X of N files complete" above convert button
  - Per-file progress bar below overall bar
  - Live status line: "filename.pdf — page 3 of 12"

- [ ] **6.5** Completion summary
  - "3 files converted. 35 slides created." replaces status line
  - Auto-open output folder if setting is on

- [ ] **6.6** Poppler missing on startup
  - `main.js` runs preflight silently on startup
  - If Poppler fails, show blocking error dialog before window loads

- [ ] **6.7** Python subprocess crash
  - `onConversionSpawnError` → show error in queue + reset button

- [ ] **6.8** Disk full / permission error
  - Python reports via `error` IPC message — handled by 3.5
  - Output folder writable check before batch — handled by 3.3

- [ ] **6.9** Resize behavior
  - Audit layout at 640×440 minimum: fix any overflow or broken elements
  - Left pane width must flex gracefully — consider reducing from fixed 420px below ~720px window width
  - Preview canvas scales with sidebar width (already `max-width: 320px; width: 100%` — verify)
  - Queue list scrollable at all heights
  - Settings modal: verify it doesn't overflow at min window height

- [ ] **6.10** Window title
  - "SlideFluid 3.0" in titlebar (fun skin: "SlideFluid ✦ 3.0" — already done in Phase 5)

- [ ] **6.11** `fileType` field on QueueItem — DOCX scaffolding
  - Add `fileType: 'pdf'` to every item created in `addFiles()`
  - No logic changes — field just needs to exist for Phase 8

- [ ] **6.12** Sidebar controls dispatch — DOCX scaffolding
  - Refactor `updateSidebar()` to call `renderSidebarControls(item)` which switches on `item.fileType`
  - `'pdf'` branch: current AR badge + fill mode selector (no change)
  - `'docx'` branch: stub comment only — font size slider goes here in Phase 8

- [ ] **6.13** Preview canvas dispatch — DOCX scaffolding
  - Refactor `drawPreview(item)` to dispatch on `item.fileType` and `settings.previewMode`
  - `'pdf'` + `'graphical'`: current canvas drawing (no change)
  - `'pdf'` + `'live'`: stub comment — PDF.js renders here in Phase 8
  - `'docx'` + either: stub comment — text simulation goes here in Phase 8

- [ ] **6.14** Preview mode setting
  - Add `previewMode: 'graphical'` to settings schema and `main.js` defaults
  - Add toggle to Appearance tab: "Preview" — Graphical (fast) / Live (renders page 1)
  - `drawPreview()` checks `state.settings.previewMode` — currently always takes graphical path

---

## Phase 7 — Packaging (separate session)

**Goal:** Distributable .dmg (Mac) and .exe (Windows).

### Tasks

- [ ] **7.1** PyInstaller — build `backend/dist/slidefluid_convert` binary
  ```bash
  cd backend
  pyinstaller --onefile --name slidefluid_convert slidefluid_convert.py
  ```

- [ ] **7.2** Bundle Poppler binaries
  - Mac (arm64 + x64): download prebuilt from `osxpoppler` or build from source
  - Windows (x64): download from `poppler-windows` releases
  - Place in `vendor/poppler/mac/` and `vendor/poppler/win/`

- [ ] **7.3** App icons
  - `assets/icons/icon.icns` (Mac) — 1024×1024 source, generate with iconutil
  - `assets/icons/icon.ico` (Windows) — multi-size .ico file

- [ ] **7.4** Build Mac .dmg
  ```bash
  npm run build:mac
  ```
  - Test on Intel Mac and Apple Silicon

- [ ] **7.5** Build Windows .exe (cross-compile or on Windows VM)
  ```bash
  npm run build:win
  ```

- [ ] **7.6** QA checklist
  - Fresh install on Mac (no Python, no Homebrew)
  - Fresh install on Windows 10 (no Python)
  - Drop zone accepts files from Finder sidebar (spec section 9)
  - Preflight all green after fresh install
  - Both skins functional
  - Cancel mid-conversion cleans up partial file

---

## Phase 8 — DOCX / TXT → PPTX ✅ COMPLETE

**Goal:** Accept Word documents and plain text files; convert text content into slides using blank-line boundaries.

### Decisions made in scoping interview
- **Formats**: `.docx` and `.txt` accepted. `.doc` rejected with friendly notice ("re-save as .docx or .txt").
- **Slide boundaries**: Two consecutive blank lines (or two blank paragraphs in DOCX). Single blank line = paragraph break within the same slide.
- **Layout**: Centered text, auto-fit font size per slide (12–120pt range), no forced title/body split.
- **Formatting preserved**: bold, italic, underline, explicit RGB font colors. Theme colors skipped (not portable).
- **Images/tables**: Skipped with a warning badge on the queue item.
- **Overflow floor**: 12pt minimum; overflow triggers a warn IPC message.
- **`.doc` binary format**: Not supported — rejected at drop with a notice.

### What was built

**`backend/slidefluid_convert.py`**
- `_parse_txt(path)` — splits on `\n\n\n+`; normalises `\r\n` and `\xa0` (non-breaking spaces from Word/Google Docs exports)
- `_parse_docx(path)` — splits on 2+ consecutive blank paragraphs; captures bold, italic, underline, explicit RGB color per run; skips images (`<w:drawing>`) and tables with warnings
- `_estimate_fits()` / `_fit_font_size()` — binary search 12–120pt; short slides scale up, dense slides shrink down
- `_add_text_slide()` — centered text box, word wrap, all formatting applied via python-pptx `RGBColor`
- `docx_info(path)` — fast slide/word count; emits `{"type":"docx_info",...}` in IPC mode
- `convert_text_doc()` — same IPC event schema as PDF pipeline (start/progress/done/error/batch_done)
- `collect_files()` — replaces `collect_pdfs`; folders now expand to .pdf + .docx + .txt
- `run_batch()` — routes on file extension; DOCX/TXT files bypass fill mode / DPI / Poppler

**`src/main.js`**
- `docx:info` IPC handler (spawns Python with `--docx-info <path>`)
- `dialog:openFiles` now offers PDF, DOCX, TXT filters
- `pdf:scan` now returns all supported types from folder expansion

**`src/preload.js`**
- `getDocxInfo(filePath)` exposed

**`src/renderer/app.js`**
- `addFiles()` routes `.docx`/`.txt` → `fileType:'docx'` items; calls `getDocxInfo` on drop
- `.doc` → `showDocNotice()` (friendly error, 6s auto-dismiss)
- First DOCX/TXT drop shows `showDocxGuidance()` banner explaining the two-blank-line boundary rule (10s, teal-tinted)
- `renderSidebarControls()` docx branch: slide count + word count badge via `renderDocxInfo()`
- `renderFileDetails()` docx branch: Slides / Words rows instead of Pages / Aspect
- `drawPreview()` docx branch: white slide canvas with file type and slide count
- `beginConversion()`: stem strips `.docx`/`.txt`; DOCX items grouped separately after PDF fill-mode groups

### Known limitations
- Theme colors (e.g. text styled with a document theme palette) are not transferred — only explicit RGB colors carry through
- Highlight colors are not transferred (pptx has no direct equivalent)
- Images and tables are skipped silently with a warning badge
- Font family is not preserved (uses PowerPoint default — Calibri)

---

## IPC contract reference

### Python → Electron (stdout JSON lines)
```json
{"type": "start",      "file": "...", "total_files": N, "file_index": N}
{"type": "progress",   "file": "...", "page": N, "total_pages": N, "message": "..."}
{"type": "done",       "file": "...", "output": "...", "slides": N}
{"type": "error",      "file": "...", "message": "..."}
{"type": "batch_done", "converted": N, "skipped": N, "errors": N, "total_slides": N}
{"type": "warn",       "file": "...", "message": "..."}
{"type": "preflight_result", "results": {...}}
```

### Electron API (`window.slidefluid`)
```javascript
// Settings
getSettings()                      → Promise<SettingsObj>
setSetting(key, value)             → Promise<true>
setSettings(obj)                   → Promise<true>

// Dialogs
openFilePicker()                   → Promise<string[]>
openFolderPicker(defaultPath?)     → Promise<string|null>
openFolder(folderPath)             → Promise<boolean>

// Conversion
startConversion(payload)           → Promise<{ok, error?}>
cancelConversion()                 → Promise<boolean>
notifyJobDone()                    → void
checkOverwrite(filePath)           → Promise<'ok'|'overwrite'|'skip'|'rename'|'cancel'>

// Events (return unsubscribe fn)
onConversionMessage(cb)
onConversionStderr(cb)
onConversionExit(cb)
onConversionSpawnError(cb)

// Preflight
runPreflight()                     → Promise<true>
onPreflightResult(cb)

// App info
getVersion()                       → Promise<string>
getPlatform()                      → Promise<{platform, arch, version}>

// To add in Phase 1:
getPdfInfo(filePath)               → Promise<{ok, pageCount?, ar?, widthPt?, heightPt?}>
scanPaths(paths[])                 → Promise<string[]>   // expands folders to PDF list
```

---

## Settings schema
```javascript
{
  outputDir: null,            // string|null
  dpi: 72,                    // 72|144
  fillMode: 'black',          // 'black'|'color_match'|'smear'
  autoOpenOnComplete: true,   // boolean
  overwriteBehavior: 'ask',   // 'ask'|'overwrite'|'skip'|'rename'
  filenameSuffix: '',         // string
  skin: 'professional',       // 'professional'|'fun'
  previewMode: 'graphical',   // 'graphical'|'live'  (live = PDF.js render, Phase 8)
}
```

---

## Context window strategy

Each phase is designed to fit in one session. Start a session by referencing this file and the relevant existing files. Phases 1–3 are the critical path (working conversion end-to-end). Phases 4–5 add settings and the fun skin. Phase 6 is polish. Phase 7 is packaging.

**Session startup prompt template:**
> I'm building SlideFluid 3.0. Read PLAN.md, then read the files relevant to the current phase. We're working on Phase N — [goal].
