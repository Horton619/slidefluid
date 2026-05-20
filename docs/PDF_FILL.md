# PDF → PPTX (rasterize + pillarbox fill)

> Load this doc when touching: `convert_pdf`, `add_page_to_pptx`, `fill_black`, `fill_color_match`, `fill_smear`, `_bar_dims`, `detect_ar`, `is_native_169`, `resolvePopplerPath`, vendored Poppler binaries, anything in `vendor/poppler/`, DPI defaults, the renderer's fill-mode picker. Trigger keywords: PDF, raster, fill mode, pillarbox, color match, smear, Poppler, pdftoppm, pdfinfo, DPI.

## TL;DR

PDFs become PPTX decks by rasterizing each page (`pdf2image` → Poppler `pdftoppm`), then pillarboxing non-16:9 pages to fit a canonical 13.33×7.5 in slide. Three fill modes for the bars: solid black, color-matched (sample edges + average), or smear (blur-extend the edge pixels outward). Poppler is vendored per-platform under `vendor/poppler/{mac,win,linux}/` and the resolver in `main.js` knows the Windows quirk (binaries live in a `bin/` subdir there only). DPI default is 144 — 72 was pixelated on text-heavy slides; anything above 144 inflates file size without visible quality gain.

## The decisions / invariants

- **Target slide aspect = 13.33 × 7.5 in (16:9).** Constant `TARGET_AR` = 16/9. Non-16:9 source pages get pillarboxed horizontally to fit. Vertical letterboxing (tall sources) is out of scope — was deliberately not implemented because real production PDFs are always wider than 16:9, never taller.
- **DPI default is 144.** 72 produced pixelated text on text-heavy slides; 144 is the sweet spot. Anything higher inflates DMG size without visible improvement to a projected slide.
- **Fill mode chain falls back on error:** `fill_smear` → `fill_color_match` → `fill_black`. Each catches its own exceptions and returns the next-simpler mode's output. The user never sees a conversion fail because of a fill-mode quirk; worst case is a black bar where they expected a smear.
- **Color match samples the outermost 5px on each side**, averaged. Capped at `width // 4` so very narrow pages don't sample the whole interior. Sampled from both edges, concatenated, averaged into one flat color. Both bars get the same color — intentional, asymmetric bars look wrong even when source content is asymmetric.
- **Smear repeats the edge column outward then Gaussian-blurs** the bar region only. Blur radius is `bar_w // 3 + 4` — calibrated so narrow bars get tighter blur (preserves edge detail) and wide bars get a soft gradient. The middle (actual page content) is pasted unblurred on top.
- **Per-batch fill mode, not per-page.** The CLI takes one `--fill black|color_match|smear` value. The renderer groups queued items by `fillMode` and spawns one backend invocation per group. If you want a mixed-mode batch, the renderer is the place to split.
- **Poppler is vendored per platform** under `vendor/poppler/{mac,win,linux}/`. Mac bundles dylibs in `libs/` and uses `@executable_path/libs/` rpaths (set via `dylibbundler` — see global CLAUDE). Windows binaries live in a `bin/` subdir of the platform folder; `resolvePopplerPath` knows this.
- **In dev:** `resolvePopplerPath` prefers the vendored binaries but falls back to system Poppler from PATH (return `null` = let pdf2image use the system). In packaged builds it's always vendored — Windows users won't have Poppler installed.
- **`pdf2image` requires `poppler_path`** to be the directory containing `pdftoppm`/`pdfinfo`, not the path to either binary. `main.js` does the platform-specific resolution and passes it to the backend via `--poppler-path`.

## Code references

| File | What it owns |
|---|---|
| `backend/slidefluid_convert.py` `convert_pdf` | Top-level PDF→PPTX entry. Loops pages via `pdf2image.convert_from_path`, applies fill, calls `add_page_to_pptx`. |
| `backend/slidefluid_convert.py` `add_page_to_pptx` | Adds one rendered page-image as a slide-sized picture in a blank PPTX slide. |
| `backend/slidefluid_convert.py` `fill_black` / `fill_color_match` / `fill_smear` | The three fill-mode implementations. |
| `backend/slidefluid_convert.py` `FILL_FUNCS` | `dict` mapping `--fill` value → function. |
| `backend/slidefluid_convert.py` `_bar_dims` | Shared math: given a non-16:9 source image, returns `(canvas_w, img_w_scaled, bar_w)`. Used by all three fill modes. |
| `backend/slidefluid_convert.py` `detect_ar` / `is_native_169` | Aspect-ratio detection. Native-16:9 pages bypass fill entirely. |
| `backend/slidefluid_convert.py` `TARGET_AR`, `SLIDE_WIDTH_IN`, `SLIDE_HEIGHT_IN`, `SLIDE_WIDTH_EMU`, `SLIDE_HEIGHT_EMU` | Canvas constants. EMUs are what python-pptx uses; inches are what humans use. |
| `src/main.js` `resolvePopplerPath` | Per-platform Poppler dir resolver. Dev: vendor → system PATH. Packaged: bundled. Windows quirk: `bin/` subdir. |
| `vendor/poppler/mac/` | `pdfinfo` + `pdftoppm` + bundled dylibs in `libs/`, rpath'd to `@executable_path/libs/`. |
| `vendor/poppler/win/bin/` | `pdfinfo.exe` + `pdftoppm.exe` + required DLLs. CI workflow downloads from oschwartz10612/poppler-windows on Windows runner. |

## What NOT to do

- ❌ **Don't change DPI default below 144** without retesting text-heavy slides. 72 looked fine in vector previews but anti-aliased text rendered through `pdftoppm` looked pixelated on a projector at 72. 144 fixed it; was a real user complaint pre-fix.
- ❌ **Don't remove the `try/except` fallback chain** in the fill modes. Each catches its own failures (numpy edge cases, weird pixel formats) and falls back to a simpler mode. Without the chain a single weird PDF page crashes the whole batch.
- ❌ **Don't sample more than 5px for color_match.** Larger samples pull in content from the page edge (logos, watermarks, footers) and the average color drifts off the actual border color.
- ❌ **Don't add vertical letterboxing** for tall source pages without a real ask. Was considered, deferred. Real-world source PDFs are always wider than 16:9; tall pages are a content-creator bug we don't try to paper over.
- ❌ **Don't pass `pdftoppm` directly as `poppler_path`.** `pdf2image` wants the *directory* containing the binary. Symptom of getting it wrong: pdf2image silently falls back to system PATH and works in dev, fails on a CI runner that has no system Poppler.
- ❌ **Don't refactor `resolvePopplerPath` to drop the Windows `bin/` subdirectory branch.** Mac/Linux Poppler binaries live at `<platform>/`; Windows lives at `<platform>/bin/`. The branch is necessary because oschwartz10612's Windows build packages with a `bin/` directory and we mirror their layout.
- ❌ **Don't change the dylibbundler invocation in the Mac packaging recipe without explicit `-s /opt/homebrew/lib/ -s /opt/homebrew/opt/poppler/lib/` search paths.** Homebrew arm64 binaries use `@rpath` instead of absolute paths; without the search paths dylibbundler hangs waiting for interactive input. (See global CLAUDE if it ever needs re-bundling.)
