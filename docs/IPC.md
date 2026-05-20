# IPC: backend ↔ main ↔ renderer protocol

> Load this doc when touching: `conversion:start`, `conversion:message`, the backend's `_emit` calls, `ConversionJob`, `_handleMessage`, `_buildArgs`, `onConversionMessage`, `beginConversion` (the renderer's group-and-spawn loop), `handleConversionMessage`, or anything in `src/preload.js`. Trigger keywords: IPC, subprocess, NDJSON, batch, group, fillMode, slideTheme, textAlign, pptxMode, splitChoices, _emit.

## TL;DR

A conversion batch is a Python subprocess spawned by `ConversionJob` in main.js. The Python side emits one JSON object per stdout line; `ConversionJob._handleMessage` parses each line and forwards it to the renderer on the `conversion:message` channel. Queue items in the renderer have a `fileType` (`pdf` | `docx` | `pptx`) and per-type config — the `beginConversion` loop groups items by their config (different group keys per fileType) so each group can be one spawn invocation. The IPC schema is defined in THREE places implicitly: backend `_emit` call sites, main.js `_handleMessage`, renderer `handleConversionMessage`. This doc is the source of truth.

## The protocol (NDJSON over stdout)

Every line of the backend's stdout in `--ipc` mode is a single JSON object with a `type` field. Recognized types:

| `type` | Payload fields | Emitted when |
|---|---|---|
| `start` | `file`, `file_index` (1-based), `total_files` | A file enters processing. |
| `progress` | `file`, `page` (1-based), `total_pages`, `message` | After each page/slide processed. |
| `done` | `file`, `output`, `slides`, optional `outputs[]`, optional `edits[]`, optional `pptx_mode` | A file finished successfully. |
| `warn` | `file`, `message` | Non-fatal issue (overflow at min font, missing notes, etc.). |
| `error` | `file`, `message` | A file failed. Batch continues to next file. |
| `batch_done` | `converted`, `skipped`, `errors`, `total_slides` | Final summary; one per spawn invocation. |
| `pptx_info` | `slideCount`, `totalNotesChars`, `slidesWithNotes`, `totalBuilds`, `slidesWithBuilds` | One-shot probe response (`--pptx-info <path>`). |
| `docx_info` | `ok`, `slideCount`, `wordCount`, or `ok: false`, `message` | One-shot probe response (`--docx-info <path>`). |
| `docx_analyze` | `ok`, `total_slides`, `overflow_slides[]` (with `slide_index`, `heading`, `lines`) | One-shot probe response (`--analyze <path>`). See `docs/DOCX_TXT.md`. |
| `preflight_result` | `results: { check_name: {ok, message} }` | Preflight (`--preflight`) finished. |

### Back-compat invariants

- **`done.output` is the primary output path (legacy single-output).** New consumers should prefer `outputs[]`, but `output` must stay populated for anything that produces a single file. Companion mode populates BOTH: `outputs: ["audience.pptx", "companion.pptx"]` AND `output: "audience.pptx"`.
- **`edits[]` is optional.** Present only when slides changed shape (split, duplicated). Per-source-slide entries like `{src_slide: 24, kind: "duplicated", chunk_count: 2, result_slides: [29, 30]}`. The renderer feeds this into the Conversion Details modal.
- **`pptx_mode` is optional.** Echo of the input mode (`teleprompter` | `split` | `companion`). Renderer uses it for the right post-conversion banner.

## Spawning a batch: how the three pieces line up

```
RENDERER (app.js)                MAIN (main.js)                  BACKEND (slidefluid_convert.py)
───────────────────              ──────────────                  ──────────────────────────────
beginConversion()
  ↓ groups items by per-type
    config (see below)
  ↓ for each group:
window.slidefluid                conversion:start ipc handler
  .startConversion(payload) ──→  ↓ instantiates ConversionJob
                                 ConversionJob._buildArgs()
                                 ↓ assembles CLI flags
                                 spawn(backendBin, [...args]) ──→  argparse → run_batch()
                                                                  ↓ per file:
                                 _handleMessage(line)              _emit({type: 'start', ...})  ←──
  onConversionMessage(payload) ←─ ↓ webContents.send             _emit({type: 'progress', ...})
                                   'conversion:message'           _emit({type: 'done', ...})
                                                                  _emit({type: 'batch_done', ...})
  ↓ batch_done resolves the
    await; loop to next group
```

## Group keys (renderer side)

`beginConversion` partitions waiting items into groups before spawning. Each group → one backend invocation. Group key shape varies by fileType:

| `fileType` | Group key | CLI flags it produces |
|---|---|---|
| `pdf` | `fillMode` | `--fill black|color_match|smear` |
| `docx` | `slideTheme \| textAlign` | `--slide-theme light|dark`, `--text-align left|center` |
| `pptx` | `pptxMode \| pptxThreshold \| pptxIncludeHeaders \| pptxIncludeSlideNumbers \| pptxTheme` | `--pptx-mode`, `--pptx-threshold`, etc. |

Different fileTypes can't share a group (the backend invocation has one input pipeline). Within a fileType, items with identical config share a group.

## CLI surface

The backend has multiple modes selected by mutually-exclusive flags:

| Mode | Trigger flag(s) | Produces |
|---|---|---|
| Batch conversion | files as positional args | NDJSON event stream |
| Preflight | `--preflight` | One `preflight_result` event |
| Docx-info probe | `--docx-info <path>` | One `docx_info` event |
| Pptx-info probe | `--pptx-info <path>` | One `pptx_info` event |
| Docx-analyze probe | `--analyze <path>` | One `docx_analyze` event |

All modes require `--ipc` (JSON over stdout) when spawned by the app. Without it the backend uses human-readable output.

Per-pipeline flags (passed through by `ConversionJob._buildArgs`):
- PDF: `--dpi`, `--fill`, `--poppler-path` (optional)
- DOCX/TXT: `--slide-theme`, `--text-align`, `--split-choices '<JSON>'`
- PPTX: `--pptx-mode`, `--pptx-threshold`, `--pptx-include-headers`, `--pptx-include-slide-numbers`, `--pptx-theme`
- Universal: `--output-dir`, `--suffix`, `--overwrite`, `--ipc`

## Code references

| File | What it owns |
|---|---|
| `backend/slidefluid_convert.py` `_emit` | The single function that writes JSON to stdout. ALL backend → renderer events flow through this. |
| `backend/slidefluid_convert.py` `_ipc_mode` | Module-level flag; set by `--ipc`. Toggles `_emit` between JSON and human-readable. |
| `backend/slidefluid_convert.py` `run_batch` | Top-level batch loop. Picks per-file converter, emits `start`/`done`/`error` per file, `batch_done` at end. |
| `backend/slidefluid_convert.py` `convert_pdf` / `convert_text_doc` / PPTX-mode functions | Per-pipeline emitters of `progress`/`warn` events. |
| `src/main.js` `class ConversionJob` | Subprocess wrapper. Owns one Python child process. |
| `src/main.js` `ConversionJob._buildArgs` | Maps the JS payload to CLI flags. **The single point of truth for arg names.** |
| `src/main.js` `ConversionJob._handleMessage` | Parses one NDJSON line, mirrors `error`/`warn` to the app log, forwards to renderer. |
| `src/main.js` `ConversionJob.start` / `cancel` | Lifecycle. `cancel` sends SIGTERM. |
| `src/main.js` IPC handlers (line ~533 onward) | All `ipcMain.handle(...)` declarations. 26 channels. |
| `src/preload.js` | The contextBridge surface — what the renderer can call. Anything the renderer needs goes through here as a one-line wrapper. |
| `src/renderer/app.js` `beginConversion` | Groups items, awaits each group's `batch_done` before spawning the next. |
| `src/renderer/app.js` `handleConversionMessage` | Renderer's IPC dispatcher. Per-`type` branches. |

## What NOT to do

- ❌ **Don't add error handling around internal-only IPC calls.** Validate at the boundaries (file load, GitHub API, user input) and trust internally. The `error` event from the backend is for user-facing failures, not for paranoid wrapping.
- ❌ **Don't drop `output` from the `done` payload when adding multi-output modes.** Always populate both `output` (primary, single path) and `outputs[]` (full list). Old consumers checking `output` mustn't break when companion mode runs.
- ❌ **Don't rename CLI flags without updating ALL THREE sites:** the backend's `argparse` definitions, `ConversionJob._buildArgs`, and `docs/DOCX_TXT.md` / `docs/PPTX_PIPELINE.md` references. The flag names ARE the contract.
- ❌ **Don't add a new IPC `type` value without updating both `_handleMessage` and `handleConversionMessage`.** Unknown types get silently dropped on the renderer side — symptoms appear as "the feature half-works."
- ❌ **Don't mix fileTypes in a single backend spawn.** Each spawn is one pipeline. Group on the renderer side.
- ❌ **Don't `print()` from the backend in IPC mode.** Use `_emit` only. Any stray stdout-print breaks NDJSON parsing on the next line.
- ❌ **Don't introduce a non-stdout side channel** (e.g., write to a file the renderer polls). The single NDJSON stream is the contract; back-and-forth lives in `--pptx-info`-style one-shot probes, not running conversations.
- ❌ **Don't widen `preload.js` casually.** Every method on `window.slidefluid` is part of the API surface. Keep it narrow; bundle related calls into a single IPC handler if you find yourself adding several at once.
