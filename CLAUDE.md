# SlideFluid — working agreement

> **Project-scoped CLAUDE.** Loads when working in `~/slidefluid`.
> Cross-project lessons live in `~/.claude/CLAUDE.md` (the "global CLAUDE").

This file is for Claude. It is **not** project documentation — it is a
collaboration contract between Dave and whatever session is reading it.
Project facts go in source code and commit messages; this file captures
the non-obvious decisions and the canonical context for working on this
codebase.

---

## Project state

SlideFluid is a Mac/Windows Electron desktop app that converts files into
PPTX presentations. **As of v3.2.0** it supports three input types:

- **PDF → PPTX** — rasterized slides with pillarbox fill modes (black,
  color-match, smear)
- **DOCX / TXT → PPTX** — text content rendered as full slides with
  auto-fit font sizing
- **PPTX → notes pipeline** — extract speaker notes into a teleprompter
  `.txt`, split long-note slides into duplicates with chunked notes
  (Split Deck), or produce a paired audience+companion deck for two-machine
  click-synced presentation (Companion)

Internal distribution for AVSC / Visual Entropy Productions; signed +
notarized with the VEP-wide Developer ID (team `L5KZ5KGKXC`). See the
global CLAUDE for the team-cert recipe + the five GitHub Secrets.

---

## Canonical references

Read these first, in this order, before touching anything non-trivial:

1. **`backend/slidefluid_convert.py`** — single-file Python backend.
   The PDF path, the DOCX/TXT path, and the PPTX-notes path all live
   here. ~1500 lines.
2. **`src/main.js`** — Electron main process. `ConversionJob` class,
   `SettingsStore`, `Logger`, all IPC handlers, auto-updater wiring,
   path resolution for dev vs packaged.
3. **`src/preload.js`** — narrow IPC API exposed to the renderer via
   `contextBridge`. Anything new the renderer needs to call belongs
   here as a one-liner wrapper.
4. **`src/renderer/app.js`** — the renderer. State, queue, drop zone,
   sidebar dispatcher (per-fileType branches), conversion flow.
   Largest single file in the project (~2500 lines) — read carefully.
5. **`src/renderer/index.html` + `styles.css`** — layout, modals,
   banners, CSS custom properties for the Pro / Fun skin system.

---

## Architecture overview

### Backend ↔ Renderer protocol

Backend is a Python subprocess spawned per conversion batch. It emits
newline-delimited JSON on stdout, one object per line:

```
{"type": "start", "file": "...", "file_index": N, "total_files": N}
{"type": "progress", "file": "...", "page": N, "total_pages": N, "message": "..."}
{"type": "done", "file": "...", "output": "...", "slides": N,
                 "outputs": [...], "edits": [...], "pptx_mode": "..."}
{"type": "warn", "file": "...", "message": "..."}
{"type": "error", "file": "...", "message": "..."}
{"type": "batch_done", "converted": N, "skipped": N, "errors": N, "total_slides": N}
{"type": "pptx_info", ...}
{"type": "docx_info", ...}
{"type": "docx_analyze", ...}
{"type": "preflight_result", "results": {...}}
```

- `output` is the primary output path (back-compat).
- `outputs: [...]` is set when a single source produced multiple files
  (Companion mode → audience + companion).
- `edits` lists per-source-slide changes for the Conversion Details modal
  (e.g. `{src_slide: 24, kind: "duplicated", chunk_count: 2, result_slides: [29, 30]}`).

### State shapes (renderer)

Queue items vary by `fileType`:

- **`pdf`** — pageCount, ar, widthPt, heightPt, fillMode
- **`docx`** — slideCount, wordCount, slideTheme, textAlign, splitChoices,
  overflowSlides
- **`pptx`** — slideCount, totalNotesChars, slidesWithNotes, totalBuilds,
  slidesWithBuilds, pptxMode (`teleprompter`|`split`|`companion`),
  pptxThreshold (effective lines), pptxIncludeHeaders,
  pptxIncludeSlideNumbers, pptxTheme (`light`|`dark`), outputPaths,
  edits, warnings, pptxModeUsed

Shared across types: id, path, name, status, progress, progressMsg,
outputPath, errorMsg.

`renderSidebarControls(item)` dispatches by `item.fileType` — that's the
extension point for adding new input types.

`beginConversion()` groups waiting items by their conversion config
(PDFs by `fillMode`; DOCX/TXT by `slideTheme|textAlign`; PPTX by
`mode|threshold|headers|slidenums|theme`). Each group becomes one
backend invocation; the renderer awaits batch_done between groups.

---

## PPTX → notes pipeline (the v3.2.0 feature)

Drop a `.pptx`, pick one of three modes. Output naming is hardcoded
per-mode (the user's global filename-suffix is *appended* after the
mode tag):

| Mode | Output | What it does |
|------|--------|--------------|
| `teleprompter` | `{stem}_notes{suffix}.txt` | Speaker notes lifted out per slide, optionally with `[Slide N]` headers. Long notes split into multiple sections. |
| `split` | `{stem}_split{suffix}.pptx` | Source slides with long notes get duplicated. First instance keeps animations, clones strip them. Each chunk goes into its duplicate's notes pane with `[click to continue]` on non-final chunks. |
| `companion` | `{stem}_audience{suffix}.pptx` + `{stem}_companion{suffix}.pptx` | Paired decks click-synced for two-machine presentation. Audience deck = same as Split. Companion deck = chunks become full-screen text slides via the existing `_add_text_slide` engine, with `chunks[0] × (builds + 1)` expansion so click count matches audience exactly through all build animations. |

### Chunking

Threshold is **effective lines**, not characters — character count gives
4-char stub fragments for slides with `"Joel\n\n[long body]\n\n[short ending]"`
structure. Each `\n` counts as one line; long lines wrap at ~65 chars
per visible line; blank lines count as gaps.

```python
_PPTX_NOTES_CHARS_PER_LINE = 65   # presenter-view notes pane width estimate
_PPTX_SOFT_BUFFER = 1.2           # don't split if effective_lines <= threshold * 1.2
_PPTX_MARKER_RE = re.compile(r"^-{3,}\s*$")   # override marker
```

Default threshold = 12 lines (persisted to settings, applied to all
queued pptx items, future drops pick up user's value).

- **Override marker `---`** on its own line is **always honored** as a
  forced chunk boundary. Threshold logic runs *within* each segment
  between markers.
- **Greedy packer merges chunks shorter than half-threshold** into a
  neighbor — speaker-name stubs ("Joel", "Matt") and short closing lines
  don't get their own slide.
- **Paragraph-level packing falls through to sentence-level** when any
  one paragraph alone exceeds threshold (avoids the `[4, 769, 12]` chunk
  pattern).

### Slide cloning (the deep-copy recipe)

`_pptx_clone_slide(prs, src_slide, keep_animations)` builds a clone via:

1. `prs.slides.add_slide(src_slide.slide_layout)` — start with the same
   layout (preserves master/theme chain).
2. Strip layout-injected placeholder shapes from the new spTree.
3. Build `rid_map: src_rId → new_rId` for source's rels. For each rel:
   - **External rels** (URL hyperlinks): `get_or_add_ext_rel(reltype, target_ref)`
     — `target_part` is undefined for external rels and will raise.
   - **Owned reltypes** (chart, diagram bundle, OLE — see below):
     `_pptx_deep_copy_part` recursively duplicates the target part so the
     clone owns its own copies.
   - **Everything else** (images, theme, layout): shared via
     `get_or_add(reltype, target_part)` — correct PPTX behavior.
4. Deep-copy each shape XML from source's spTree, rewriting `r:*`
   attributes per `rid_map`.
5. Carry over `<p:cSld>/<p:bg>` (slide-level background). **Title slides
   with no shapes store their entire visual here — without this copy
   clones render blank.**
6. Carry over `<p:clrMapOvr>` (color map override) and `<p:transition>`.
7. If `keep_animations=True`, carry over `<p:timing>` (with `rid_map`
   applied to its rId references).

**Owned reltypes** — PowerPoint expects single-slide ownership of these,
so sharing crashes:

```python
_PPTX_OWNED_RELTYPES = frozenset({
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + t
    for t in ("chart", "diagramData", "diagramLayout", "diagramColors",
              "diagramQuickStyle", "diagramDrawing", "oleObject", "package")
})
```

If you add a new reltype that needs single-slide ownership (e.g.
embedded video w/ ownership semantics, if PowerPoint ever complains
about shared video refs), drop it here and `_pptx_deep_copy_part`
handles the recursion automatically.

### python-pptx 1.0 API gotchas

These bit me; preserved for the next agent:

- **`Part.load(partname, content_type, package, blob)`** — *package
  comes before blob*, not after. Older docs and many gists have blob
  first; that's wrong in 1.0 and the symptom is `TypeError: object of
  type 'Package' has no len()` deep inside zip serialization.
- **`Package` in 1.0 has no `_parts` dict.** Parts are discovered via
  the rels graph (`iter_parts()` walks from the root). New parts get
  persisted on save when they're reachable through some chain of rels
  — no manual `package._parts[uri] = new_part` needed (and that would
  raise `AttributeError` anyway).
- **`_Relationships.get_or_add(reltype, target_part)` returns the rId
  string**, not a `_Relationship` object. `new_rel.rId` raises
  `AttributeError: 'str' object has no attribute 'rId'`.
- **External rels need `get_or_add_ext_rel`** — calling `.target_part`
  on an external rel raises "target_part property on _Relationship is
  undefined when target-mode is external".
- **Recursive deep-copy needs a `_reserved` URI set.** Parts created in
  the same recursion aren't yet wired into the rel graph, so
  `iter_parts()` doesn't see them when computing the next free URI.
  Without `_reserved`, you get colliding partnames.

### File-write atomicity

All conversion outputs go through a temp file + `os.replace`:

```python
def _atomic_save_pptx(prs, out_path):
    tmp = out_path.with_name(out_path.name + ".tmp")
    prs.save(str(tmp))
    os.replace(str(tmp), str(out_path))
```

Without this, opening an 868 MB PPTX mid-save (~10s of writing) throws
PowerPoint errors that aren't actually bugs in our output.

### IPC additions

- `pptx:info` in main.js → calls backend `--pptx-info <path>`,
  returns `{slideCount, totalNotesChars, slidesWithNotes, totalBuilds, slidesWithBuilds}`.
- `done` payload extended with optional `outputs: [...]` (for Companion)
  and `edits: [...]` (for the Conversion Details modal).
- Backend `_emit` mirror to log on error/warn (in `ConversionJob._handleMessage`)
  so postmortem diagnosis from `~/Library/Application Support/slidefluid/slidefluid.log`
  is sufficient.

---

## Decisions already made (don't relitigate)

These landed after iteration. New information is welcome; revisiting
without new information wastes turns.

- **Threshold is lines, not chars.** Character count gives 4-char stub
  fragments for natural-prose slides.
- **Default = 12 effective lines.** Roughly matches Presenter View's
  visible notes-pane height without scrolling. User changes persist.
- **Soft buffer 1.2×.** A 13-line note shouldn't split; 14 might; 15+
  definitely.
- **`---` override marker is always active**, not a toggle. UX hint
  lives in the popout, not the sidebar.
- **Companion deck uses canonical 16:9 dimensions** (13.33 × 7.5 in),
  not source aspect. Companion is for a confidence monitor, not the
  audience screen.
- **Animations preserved on first instance only.** Cloned duplicates of
  a slide with click-builds get animations stripped. Companion deck
  uses `chunks[0] × (builds + 1)` expansion to stay click-synced
  through all build animations.
- **Notes-slide rels are skipped during clone** — each clone gets its
  own fresh notes pane populated explicitly by the caller. Don't carry
  over the source's notes_slide.
- **Apply-to-all lives inside the Split Settings popout**, not as a
  permanent UI element. Visible only when ≥2 pptx items in queue.
- **Slide-numbers toggle is ONE UI control backing TWO different fields**
  (`pptxIncludeHeaders` for teleprompter, `pptxIncludeSlideNumbers` for
  split/companion). A hint line under it explains the mode-specific
  effect. Don't split into two toggles — the user reasoning is the same.
- **Done items get a clickable ⓘ** for the Conversion Details modal;
  Done ⚠ N for items with warnings. Both open the same modal.
- **Drop a file → auto-select it** so accidental re-conversions of
  completed items don't happen.
- **Pro tagline: "Transition to greatness"** (static). **Fun tagline:**
  pool of 7 in `FUN_TAGLINES`, one picked per app launch — includes
  "How does your PPT identify?" and matches the existing Fun-skin
  format-reassignment tone.

---

## Working agreement

### Ask before doing when

- The change spans 3+ files. Show a one-paragraph plan and the file
  list first.
- You're about to change a documented "Decision already made" above.
- The change touches `_pptx_clone_slide`, the chunking math, or the
  IPC schema.

### Just do it when

- One-file diff, mechanical change, no design decisions.
- Bug fix where the symptom + traceback name the cause.
- Cosmetic / copy tweak in already-shipped UI.

### Always re-read source when

- It's been more than ~5 turns since you last read the file you're
  editing. The renderer (`app.js`) changes shape fastest.
- You're about to assume a python-pptx signature. The library bit me
  twice (Part.load arg order, get_or_add return type) — verify.
- A summary handoff tells you what's there. Trust git log and the file
  tree, not the summary.

### Flag uncertainty

Saying *"I'm not sure if this also affects X — let me check"* is
welcomed. Confident tone on a guess is the failure mode.

---

## Anti-patterns

- Adding error handling around internal code. Validate at boundaries
  (IPC, file load, GitHub API), trust internally.
- Treating `rel.target_part` as safe without first checking
  `rel.is_external`.
- Sharing chart / SmartArt / OLE parts between cloned slides via shared
  rels (the PowerPoint crash that drove the deep-copy refactor).
- Splitting on character count when the goal is "fits on one slide".
- Manually writing to `package._parts[uri]` (the dict doesn't exist in
  python-pptx 1.0).
- Writing PPTX outputs directly to the final path (always atomic write).
- Generalizing the per-mode output names (`_notes`, `_split`,
  `_audience`, `_companion`) into a single template. They're
  deliberately distinct and serve as user-facing nomenclature.

---

## Stack & environment

- **Frontend:** Electron + vanilla HTML/CSS/JS renderer (no bundler, no
  TypeScript)
- **Backend:** Python 3 subprocess. Deps: `numpy`, `Pillow`,
  `pdf2image`, `python-pptx` (1.0.x), `python-docx`, `lxml`.
- **Always use `venv/bin/python3`** — never system python3. Install new
  deps with `venv/bin/pip`.
- **Vendored Poppler** in `vendor/poppler/{mac,win,linux}/`; `pdfinfo`
  + `pdftoppm` used only by the PDF path.
- **Packaging:** PyInstaller for the backend binary + electron-builder
  for the app shell. CI workflow `.github/workflows/release.yml` builds
  mac-arm64 (signed + notarized) and win-x64 (unsigned) on `v*.*.*`
  tag push.
- **Auto-updater:** `electron-updater` reads GitHub Releases on
  `latest-mac.yml` / `latest.yml`. Both `dmg` AND `zip` targets are
  required on Mac for in-place updates to work — see the global CLAUDE
  for the full recipe.

---

## Release flow

1. Bump `package.json` `"version"`
2. Commit on feature branch
3. `git push origin HEAD:main` (or via PR if shared)
4. `git tag vX.Y.Z && git push origin vX.Y.Z`
5. CI builds + publishes to GitHub Releases on the tag

Auto-update reads from there; users on v(X.Y.Z-1) pick up the new build
on launch (within the GitHub atom-feed cache window — ~5 min).
