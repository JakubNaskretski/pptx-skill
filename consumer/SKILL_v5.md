# pptx-skill — v5 agent contract

You are composing PowerPoint decks from a library of **structural
skeletons** (slide layouts with typed slots + constraints) and
**themes** (master + palette + fonts). You pick skeletons by
*function and fit*, fill their slots with your content, and the
consumer builds the deck on the chosen host theme.

## The flow

1. **List what's available**
   ```bash
   python reader.py list-themes
   python reader.py list-skeletons [--category data] [--has-slot table]
   ```

2. **Find skeletons that fit your content**
   ```bash
   python reader.py match-skeletons \
     --content '{"title": "Q4 results beat", "bullets": ["Revenue +12%", "Margin up"]}' \
     --category data
   ```
   Returns ranked candidates with `fit_score`, `slot_mapping`, and
   `headroom`. Higher `fit_score` = tighter fit (the layout was
   designed for that content length — loose fit wastes the layout).

3. **If zero matches: rephrase, don't pick a near-miss**
   When `match-skeletons` returns `matches: []`, each `issues[]` entry
   includes a concrete `suggested_action` ("rephrase to ≤60 chars
   (drop 13)"). Your job is to **shorten the content**, not to pick
   a skeleton whose constraints don't match. Re-call `match-skeletons`
   after rephrasing.

   **Escape hatch.** If rephrasing would lose meaning (text is
   already terse), wrap the value in the plan as
   `{"value": "...", "overflow": "shrink"}`. The build engine will
   auto-shrink the font and emit a warning to a sidecar file for
   the user to review manually. Use sparingly.

4. **Pre-flight check the full plan**
   ```bash
   python reader.py validate-plan plan.json
   ```
   Returns `{ok, errors, warnings}`. Hard errors block the build.
   `overflow:shrink` violations land as warnings, not errors.

5. **Build the deck**
   ```bash
   python reader.py compose-v5 plan.json out.pptx --theme <theme_id>
   ```
   Picks one host theme per output deck. All slides inherit the
   theme's master (brand bars, page numbers, footer), palette, and
   fonts. The skeletons are theme-free; identity comes from the host.

## Slot kinds

| Kind | Content shape | Constraints |
|---|---|---|
| `heading` | string | `max_chars`, `max_lines`, `required` |
| `paragraph` | string | `max_chars`, `max_lines`, `required` |
| `bullets` | list of strings | `max_items`, `max_chars_per_item`, `required` |
| `image` | `"asset_<id>"`, `{"asset": "asset_<id>"}`, or `"placeholder"` / `{"placeholder": true, "label": "..."}` for a labeled grey box | `aspect`, `required`, `auto_fit` |
| `table` | `{"rows": N, "cols": N, "has_header": bool, "data": [[...]]}` | `max_rows`, `max_cols`, `has_header` |
| `chart` | `{"type": "bar\|column\|line\|pie\|doughnut\|area" (+ `_stacked` / `_markers` variants), "categories": ["..."], "series": [{"name": "...", "values": [...]}]}` | `chart_type`, `max_series`, `max_categories` |
| `footer` | string | `max_chars`, `max_lines`, `auto_from_host` |

Each slot also carries `geometry` (fractional `x/y/w/h`) and a
`style` block with theme-relative tokens (`font_role: major|minor|
explicit`, `color_role: primary|accent|text_default|background`)
that get resolved against the chosen host theme at build time.

## Plan shape

```json
[
  {
    "skeleton_id": "deckA_03",
    "slots": {
      "title": "Q4 results beat consensus",
      "body": ["Revenue +12%", "Margin expanded 200bps", "FCF positive"],
      "hero": "asset_a1b2c3d4"
    }
  },
  {
    "skeleton_id": "deckA_07",
    "slots": {
      "title": {"value": "A slightly longer title", "overflow": "shrink"},
      "data_table": {
        "rows": 3, "cols": 2, "has_header": true,
        "data": [["Quarter", "Revenue"], ["Q3", "$1.2M"], ["Q4", "$1.8M"]]
      }
    }
  }
]
```

## Engine-side helpers

Don't compute character counts, aspect ratios, or EMU coordinates
yourself — call these instead:

```bash
python reader.py measure-text "Q4 results" --against deckA_03.title
python reader.py check-asset-fit asset_a1b2 deckA_03 hero
python reader.py find-asset --kind photo --feel warm --suitable-for team
```

`measure-text` returns `{chars, words, lines_est}` and, with
`--against`, the headroom for a specific slot. `check-asset-fit`
returns whether the asset fits a target image slot (aspect, kind,
resolution) plus a `suggestion` if not. `find-asset` returns a
deterministic shortlist — see "Picking images" below.

## Picking images

For every image slot, **call `find-asset` first** — do not scan
`index.json` and pick by free-text fields. The shortlist is filtered
purely on controlled-vocab tags (`kind`, `feel`, `composition`,
`suitable_for`, `scope`, `colors`), so two runs against the same
library produce the same candidates. Use `subject` / `depicts` only
to break ties among the returned shortlist.

```bash
python reader.py find-asset \
  --kind photo \
  --feel warm \
  --suitable-for team \
  --limit 5
```

Algorithm:

1. Call `find-asset` with the slot's required `kind` plus the deck's
   `feel` and the slot's intent (`suitable_for`).
2. If `matches: []`, drop the constraint named in `suggestion` and
   retry (broaden order: `--colors`, `--composition`, `--scope`,
   `--suitable-for`, `--feel`).
3. From the surviving shortlist, pick one whose `subject` / `depicts`
   fits the slide topic.
4. If you exhaust the broadening order and the slot is **not**
   required, omit the slot from the plan — the build skips it.
5. If the slot IS required and nothing fits, choose one:
   - **External source.** Use your own web tools to find an image,
     `POST /api/asset/add` with the file, get back an `asset_id`,
     and use it in the plan. Re-running `find-asset` after the upload
     will return it on the next call.
   - **Placeholder.** Pass `"placeholder"` (the literal string) as
     the asset value. The build draws a dashed grey box labeled
     `image needed: <slot_id>` and emits a warning in the sidecar so
     the user knows to swap it in by hand. Pass
     `{"placeholder": true, "label": "Customer logo here"}` for a
     custom hint label.

Don't pick assets by reading `index.json` directly — the free-text
fields drift between describe passes, so a vibes-based pick gives
different results on different runs. `find-asset` is the idempotent
selector; reserve free-text reading for tiebreaks among ≤5 already-
vetted candidates.

## Categories

Skeletons carry one or more functional categories — use these to
filter `list-skeletons` / `match-skeletons`:

`opening` (title / agenda) · `section_divider` (between sections) ·
`content` (general body) · `comparison` (2-column side-by-side) ·
`data` (table or chart heavy) · `metric` (single large stat) ·
`quote` (pull-quote, testimonial) · `closing` (Q&A, thank you,
next steps).

A skeleton can have multiple categories (a "Thank you" closing
slide that's also opening-shaped is `[opening, closing]`).

## What the agent does NOT do

- Re-render slides yourself; `compose-v5` owns slide construction.
- Pick a near-miss skeleton instead of rephrasing.
- Mix multiple host themes in one output deck (one `--theme` per
  build).
- Re-style master decorations (brand bars, page numbers) — those
  ride along with the chosen host theme.

## What's in the bundle

```
SKILL.md                        you are reading this
reader.py                       the consumer; ~2000 LOC; pure stdlib + python-pptx + yaml
requirements.txt                python-pptx, pyyaml
index.json                      summaries of every theme/skeleton/asset
themes/<id>/
  theme.yaml                    palette, fonts, decorations, master_pptx ref
  master.pptx                   host master for compose-v5 to build on
  preview.png                   blank-layout thumbnail (optional)
skeletons/<id>/
  skeleton.yaml                 slots, geometry, style, constraints, categories
  preview.png                   source-slide thumbnail (optional)
  background.png                frozen underlay (optional; freeze-as-background skeletons)
assets/<id>.<ext>               raster / SVG asset binaries
assets/<id>.yaml                asset descriptions (subject, depicts, feel, colors)
```

No network. No state. No vision required at compose time.
