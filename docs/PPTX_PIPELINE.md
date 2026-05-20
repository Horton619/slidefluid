# PPTX pipeline (clone, chunk, companion sync)

> Load this doc when touching: `_pptx_clone_slide`, `_pptx_deep_copy_part`, `_chunk_notes`, `_greedy_pack`, `_pptx_split_deck_into`, `_pptx_make_companion`, `_decorate_chunk`, `_pptx_count_click_builds`, animations, click-sync, anything dealing with a `.pptx` going IN as input. Trigger keywords: clone, deep-copy, chunk, threshold, builds, animations, companion, teleprompter, split deck.

## TL;DR

One `.pptx` dropped in produces one of three outputs depending on mode: a `.txt` of speaker notes (teleprompter), a duplicate-and-chunk `.pptx` (split deck), or a paired audience+companion deck for two-machine presentation. The duplicate-and-chunk side relies on a from-scratch slide-cloning routine that rebinds rIds and deep-copies "owned" parts (charts, SmartArt, OLE) so PowerPoint stays happy. Chunking math is in *effective lines*, not characters. The companion-mode click-sync formula `chunks[0] × (builds + 1) + chunks[1..]` is load-bearing — it's how clicks on the audience monitor stay in lockstep with the companion screen even when the source slide has click-triggered build animations.

## The decisions / invariants

- **Threshold is effective lines, not characters.** Character counts produce 4-char stub fragments on natural-prose slides (`"Joel\n\n[long body]\n\n[ending]"` chunks badly with char counts). `_effective_lines` counts `\n` as 1 line, wraps at `_PPTX_NOTES_CHARS_PER_LINE = 65`, treats blank lines as a 1-line gap. Default threshold = 12 effective lines, matching Presenter View's visible-without-scrolling notes pane.
- **Soft buffer is 1.2×.** A 13-line note doesn't split; 14 might; 15+ definitely. Lives in `_PPTX_SOFT_BUFFER`.
- **`---` on its own line is ALWAYS a forced chunk boundary.** Threshold splitting only runs *within* segments delimited by `---` markers. Not a toggle, not optional. Discoverability hint lives in the Split Settings popout, not the sidebar.
- **Greedy packer absorbs short neighbors.** After the initial pack, any chunk under half-threshold gets merged into a neighbor — keeps "Joel", "Matt", short closings from becoming near-empty slides on their own.
- **Paragraph-level packing falls through to sentence-level** when any paragraph alone exceeds threshold. Avoids the `[4, 769, 12]`-line chunk pattern from a tall paragraph splitting badly.
- **Animations preserved on first instance only.** First duplicate of a source slide carries `<p:timing>`; subsequent duplicates strip it. This is what makes the chunk sequence advance one chunk per click without re-firing builds.
- **Notes-slide rels are skipped during clone.** Each clone gets a fresh, empty notes pane populated by the caller. Don't carry `notesSlide` from the source — they share rels and create dupe-notes chaos.
- **Owned reltypes get deep-copied; everything else is shared.** Charts, SmartArt bundle (5 reltypes), OLE embeds need single-slide ownership — PowerPoint crashes if two slides share them. Everything else (images, theme, layout) shares correctly via `get_or_add(reltype, target_part)`.
- **Companion expansion formula is `chunks[0] × (builds + 1) + chunks[1..]`.** Per source slide with B click-builds and K chunks: audience side has K duplicates (first keeps animations, rest stripped); companion side has B+K text slides. The first chunk repeats B+1 times to cover every build click PLUS the final advance off slide 1. Click count matches exactly through all build animations.
- **Companion deck is canonical 16:9 (13.33 × 7.5 in)**, not source aspect. Companion is a confidence monitor, not the audience screen — should look like a teleprompter regardless of what the source deck does.
- **All saves are atomic.** `_atomic_save_pptx` writes to `<out>.tmp` and `os.replace`s into place. Without this, PowerPoint opening a partially-written 800MB file gives errors that aren't actual bugs in our output.
- **Deep-copy recursion needs a `_reserved` URI set.** Parts created in the same recursive call aren't wired into the rel graph yet, so `iter_parts()` won't see them when computing the next free partname. Without `_reserved`, you get colliding partnames inside the same recursion.

## python-pptx 1.0 API gotchas (these bit; don't re-discover)

- **`Part.load(partname, content_type, package, blob)`** — *package before blob*, not after. Older docs and gists have blob first; symptom of getting it wrong is `TypeError: object of type 'Package' has no len()` deep inside zip serialization.
- **`Package` in 1.0 has no `_parts` dict.** Parts are discovered via the rels graph (`iter_parts()` walks from root). New parts persist on save when they're reachable through any rel chain — no manual `package._parts[uri] = new_part` (and that would raise `AttributeError` anyway).
- **`_Relationships.get_or_add(reltype, target_part)` returns the rId STRING**, not a `_Relationship` object. `new_rel.rId` raises `AttributeError: 'str' object has no attribute 'rId'`.
- **External rels need `get_or_add_ext_rel`** — calling `.target_part` on an external rel raises "target_part property on _Relationship is undefined when target-mode is external". The clone routine checks `rel.is_external` before touching `target_part`.

## Code references

| File | What it owns |
|---|---|
| `backend/slidefluid_convert.py` `_pptx_clone_slide` | The clone recipe: rid_map build, rels rebind, shape deep-copy, `<p:bg>` / `<p:clrMapOvr>` / `<p:transition>` / `<p:timing>` carry-over. |
| `backend/slidefluid_convert.py` `_pptx_deep_copy_part` | Recursive deep-copy for owned reltypes (charts, SmartArt, OLE). Uses `_reserved` URI set to dodge partname collisions inside the same recursion. |
| `backend/slidefluid_convert.py` `_pptx_next_partname_like` | Next free partname after a given source pattern, respecting `_reserved`. |
| `backend/slidefluid_convert.py` `_remap_rids` | Rewrites `r:*` attributes inside a copied XML subtree using `rid_map`. |
| `backend/slidefluid_convert.py` `_PPTX_OWNED_RELTYPES` | The set of reltypes that need deep-copy. Extend if PowerPoint complains about shared resources on a new reltype. |
| `backend/slidefluid_convert.py` `_effective_lines` / `_greedy_pack` / `_split_by_threshold` / `_chunk_notes` | The chunking pipeline. |
| `backend/slidefluid_convert.py` `_PPTX_MARKER_RE` | The `---` forced-boundary regex. |
| `backend/slidefluid_convert.py` `_pptx_split_deck_into` | Split-deck mode. |
| `backend/slidefluid_convert.py` `_pptx_make_companion` | Companion-mode. Where the `chunks[0] × (builds + 1)` formula lives. |
| `backend/slidefluid_convert.py` `_pptx_count_click_builds` | Counts click-triggered build animations on a source slide. |
| `backend/slidefluid_convert.py` `_pptx_resolve_chunks` | Per-slide chunk resolution. Returns `(chunks, was_merged)` — the `was_merged` flag is always False now (kept for caller compatibility). |
| `backend/slidefluid_convert.py` `_decorate_chunk` | Adds `[Slide N — part k/K]` prefix + `[click to continue]` suffix per the user's toggle. |
| `backend/slidefluid_convert.py` `_atomic_save_pptx` / `_atomic_write_text` | Temp-suffix + `os.replace` write pattern. |

## What NOT to do

- ❌ **Don't share owned reltypes between cloned slides.** Charts, SmartArt parts, OLE embeds — single-slide ownership only. The deep-copy path exists for exactly this. Sharing them via `get_or_add(target_part)` will crash PowerPoint on open.
- ❌ **Don't access `rel.target_part` without first checking `rel.is_external`.** External rels (URL hyperlinks) don't have a target part. Branch on `is_external` first.
- ❌ **Don't carry over the source's `notesSlide` rel during clone.** Each clone gets its own fresh notes pane. The chunk text gets injected after the clone returns.
- ❌ **Don't write PPTX outputs directly to the final path.** Always temp-suffix + `os.replace`. PowerPoint opening a half-written 800MB file is the WC23 mid-conversion bug — solved once, don't reopen it.
- ❌ **Don't split on character count.** The goal is "fits on one slide visually." Lines, not chars. The `_effective_lines` machinery already handles wrapping + blank-line gaps.
- ❌ **Don't generalize the per-mode output names.** `_notes.txt`, `_split.pptx`, `_audience.pptx`, `_companion.pptx` are deliberately distinct — they're user-facing nomenclature. Don't refactor into a template.
- ❌ **Don't manually populate `package._parts[uri]`.** The dict doesn't exist in python-pptx 1.0. Parts persist via the rels graph at save time.
- ❌ **Don't tune `_PPTX_SOFT_BUFFER` below 1.2 or above 1.4 without checking the 13-line slide test case.** 1.2 is calibrated to "13 lines fits without scrolling; 14 might wrap." Outside that range, real notes either split too aggressively or overflow the Presenter View pane.
- ❌ **Don't strip animations on the first instance of a duplicated slide.** Build clicks must fire on instance 1 only; subsequent clones must NOT carry `<p:timing>` or the click-sync math breaks.
- ❌ **Don't change the companion expansion formula** without recomputing what "one click on audience = one slide on companion" means for build-heavy decks. The current `chunks[0] × (builds + 1) + chunks[1..]` formula is the only one that holds the invariant.
