# DOCX / TXT → PPTX pipeline

> Load this doc when touching: `_parse_txt`, `_parse_docx`, `_strip_rtf`, `_merge_orphan_headings`, `_estimate_fits`, `_fit_font_size`, `_add_text_slide`, `docx_analyze`, `_find_spill_break`, `_find_midpoint_break`, `_apply_split_choices`, `convert_text_doc`, the renderer's overflow modal / break picker, anything with `--text-align`, `--slide-theme`, `--split-choices`. Trigger keywords: docx, text doc, txt, rtf, font fit, overflow, split choice, slide theme, alignment.

## TL;DR

A `.docx` or `.txt` becomes a `.pptx` where every paragraph block (separated by 2+ blank lines) is one slide. Font size is auto-fit per slide via binary search between MIN/MAX bounds. `_estimate_fits` uses character-width and line-height ratios calibrated to Calibri at typical sizes. If a slide overflows even at min font, the renderer surfaces an overflow modal where the producer picks how to split (auto at overflow point, midpoint never-mid-sentence, or manual line pick) — choices are passed back to the backend as a `--split-choices` JSON arg and applied before the slides are built. RTF files masquerading as `.txt` get detected by their `{\rtf...` preamble and stripped to plain text first.

## The decisions / invariants

- **Slide boundary = 2+ consecutive blank lines** (TXT) / 2+ consecutive blank paragraphs (DOCX). Single blanks are paragraph breaks WITHIN a slide. `_merge_orphan_headings` is a safety net for inconsistent source docs: if a slide is just a single heading line, it gets merged into the next slide automatically.
- **Font fit constants are calibrated to system-ui Calibri.** `_CHAR_W_RATIO = 0.44`, `_LINE_H_RATIO = 1.20`. Tuning lower than 0.44 gives an overconfident estimate (fonts pick too large, text overflows the box). Higher pushes fonts too small. The current values were validated against the TSCRA / PAT / NCARB crew letters.
- **Font bounds: 20 ≤ size ≤ 54.** Anything below 20 is unreadable on a projector; anything above 54 looks like a poster. Binary search finds the largest size that fits.
- **DOCX bullet detection has THREE fallbacks** in order: (1) `numPr` element in paragraph XML (Word's real list mechanism), (2) style name contains "list" or "bullet", (3) leading text starts with `•/-/*/·`. Each was added because a real source doc didn't honor the previous one. Don't remove any layer.
- **python-docx RGBColor is a tuple subclass, NOT an object with `.red/.green/.blue` attributes.** Use `rgb[0]`, `rgb[1]`, `rgb[2]`. python-pptx's RGBColor has named attributes; they share the class name but are different classes. Mixing them up gives `AttributeError`.
- **`\xa0` (non-breaking space) is normalized to regular space** in both parsers. Word and Google Docs exports commonly produce lines containing only `\xa0` — without normalization those count as "non-blank" and break the 2-blank-line slide-boundary rule.
- **lxml XPath in f-strings needs Clark notation with TRIPLE braces.** `f".//{{{ns}}}tag"` is correct; `f".//{{{ns}}}tag"` looks identical but produces an invalid XPath if the braces are doubled instead of tripled. Triple braces produce a literal `{` after f-string substitution; double braces produce the namespace value bare and lxml rejects it.
- **RTF detection is by `{\rtf` preamble.** If the file (regardless of extension) starts with `{\rtf`, run `_strip_rtf` first. TextEdit on macOS saves RTF with a `.txt` extension by default — common path for users dropping notes.
- **Three split modes for overflow slides:** `spill` (binary-search the largest prefix that fits at min font, break there), `midpoint` (split near `len(paragraphs) // 2`, then walk to a sentence-ending paragraph — never end a slide mid-sentence), `manual` (renderer surfaces a line picker, user clicks where to break). Default suggestion in the picker is the midpoint result.
- **Continuation slides auto-prepend `<heading> cont.`** when the source slide started with a heading (first paragraph has `is_heading=True`). For TXT files this means: the first non-bullet line of any slide is treated as a heading.
- **Split choices are passed via CLI as JSON:** `--split-choices '[{"slide_index":N,"mode":"midpoint"},...]'`. The renderer collects them per-file, the backend's `_apply_split_choices` expands the slide list before rendering.
- **The `--text-align` and `--slide-theme` CLI flags are part of the IPC contract.** Don't rename without coordinating with main.js (`ConversionJob._buildArgs`) AND the renderer's `beginConversion` grouping. See `docs/IPC.md`.

## Code references

| File | What it owns |
|---|---|
| `backend/slidefluid_convert.py` `_parse_txt` | TXT parser. Two-blank slide boundary, bullet-char detection, heading-on-first-non-bullet. |
| `backend/slidefluid_convert.py` `_parse_docx` | DOCX parser. Three-tier bullet detection. Returns `(slides, warnings)`. |
| `backend/slidefluid_convert.py` `_strip_rtf` | Removes RTF control words / control symbols, returns plain text. Detected by `{\rtf` preamble. |
| `backend/slidefluid_convert.py` `_merge_orphan_headings` | Post-parse merge: any single-heading slide gets absorbed into the next slide. |
| `backend/slidefluid_convert.py` `_estimate_fits` / `_fit_font_size` | Char-width + line-height estimator + binary search for largest fit. `_CHAR_W_RATIO=0.44`, `_LINE_H_RATIO=1.20`. |
| `backend/slidefluid_convert.py` `_add_text_slide` | Builds one PPTX slide from a list of paragraph dicts. Handles theme (light/dark), alignment, headings, bullets, runs. |
| `backend/slidefluid_convert.py` `docx_analyze` | One-shot pre-conversion analysis. Returns overflow-slide list to the renderer for the resolution modal. Emits `docx_analyze` IPC message. |
| `backend/slidefluid_convert.py` `_find_spill_break` / `_find_midpoint_break` | Auto-split heuristics. Midpoint walks toward sentence ends — `.!?…` — never returns a break that ends mid-sentence. |
| `backend/slidefluid_convert.py` `_apply_split_choices` | Applies the `--split-choices` JSON to the parsed slide list. Inserts `<heading> cont.` slides between halves. |
| `backend/slidefluid_convert.py` `docx_info` | Slide count + word count probe. Called from `docx:info` IPC. |
| `backend/slidefluid_convert.py` `convert_text_doc` | Top-level entry for DOCX/TXT. Parses → applies split choices → fits font per slide → emits `start`/`progress`/`done`. |
| `src/renderer/app.js` `_showOverflowModal` / `_showBreakPicker` | Renderer's overflow resolution UI. Per-slide choice → `item.splitChoices`. |

## What NOT to do

- ❌ **Don't use system `python3`.** Always `venv/bin/python3` in dev; the bundled PyInstaller binary in packaged builds. System Python won't have python-docx/python-pptx/numpy/etc.
- ❌ **Don't tune `_CHAR_W_RATIO` or `_LINE_H_RATIO` without re-validating against the saved crew-letter samples** (TSCRA 26, PAT 25, NCARB ABM 25, LAS 25). Those are the real-world calibration points; a small ratio change flips multiple slides between "fits" and "overflows."
- ❌ **Don't treat `run.font.color.rgb` as having `.red` / `.green` / `.blue` attributes.** It's a tuple subclass. Use index access. (Yes, even though python-pptx's same-named class has attributes — different class.)
- ❌ **Don't split on character count.** Lines, not chars. See `docs/PPTX_PIPELINE.md` for the same rule in a different pipeline.
- ❌ **Don't write `re.compile` for `--`-style markers without anchoring to start-of-line.** The override marker is `^-{3,}\s*$` — a literal `---` mid-paragraph is not a slide break.
- ❌ **Don't lose the `numPr` bullet check.** Some DOCX files (Google Docs export especially) have list paragraphs with no `is_list_paragraph` style — only the `<w:numPr>` element exposes them as bullets.
- ❌ **Don't change the order of bullet-detection fallbacks** without testing all three. They handle different source-app quirks; each was added for a real file that the previous tier missed.
- ❌ **Don't accept a midpoint split that ends mid-sentence.** `_find_midpoint_break` deliberately walks toward sentence-ending punctuation. If you swap to a simpler "halve the list" approach, slides will end mid-thought and the teleprompter use case breaks.
- ❌ **Don't rename the CLI flags** (`--text-align`, `--slide-theme`, `--split-choices`) without updating `main.js` `ConversionJob._buildArgs` AND `docs/IPC.md`. They're part of the contract surface.
