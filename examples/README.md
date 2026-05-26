# Example decks

Two neutral-brand 16:9 builder scripts that emit ready-to-ingest
`.pptx` files. Together they cover the structural slides every deck
needs (cover / agenda / dividers / quote / closing) plus the content
slides that fill in between (header + text + chart/image/table).

The compiled `.pptx` outputs ship in this directory **and** the
predigested artifacts (theme.yaml, skeleton.yaml, asset binaries +
sidecars) ship under `authoring/workspace/` so a fresh clone has a
fully ingested corpus — no builder script, no `cli.py ingest` step
required. Run the workflow (`cli.py validate`, `cli.py build-v5`,
the web app, …) immediately after `git pull`.

To regenerate the predigested resources after editing a builder:

```bash
python3 examples/build_standard_templates.py
python3 examples/build_title_and_breaks.py
python3 authoring/cli.py remove-deck standard_templates
python3 authoring/cli.py remove-deck title_and_breaks
python3 authoring/cli.py ingest examples/standard_templates.pptx
python3 authoring/cli.py ingest examples/title_and_breaks.pptx
```

(If the asset SHAs change, update the explicit exceptions in
`.gitignore` to match the new file names.)

## Standard content templates — `build_standard_templates.py`

Ten content layouts. Each: header at top, text or bullets on the
left, chart / image / table on the right.

| #  | Left content         | Right content                  |
|----|----------------------|--------------------------------|
| 1  | 5 bullets            | Bar chart (regional revenue)   |
| 2  | 4 bullets            | Line chart (12-month trend)    |
| 3  | 4 bullets            | Pie chart (revenue mix)        |
| 4  | 4 bullets            | Image placeholder              |
| 5  | 4 bullets            | Table (5 rows × 3 cols)        |
| 6  | Paragraph            | Column chart (margin trend)    |
| 7  | Paragraph            | Image placeholder              |
| 8  | Paragraph            | Table (4 rows × 3 cols)        |
| 9  | Lead + 4 sub-bullets | Image placeholder              |
| 10 | 2 KPI stats          | Column chart (quarterly)       |

## Title, breaks, closing — `build_title_and_breaks.py`

Six structural layouts that surround the content slides.

| # | Layout                                                          |
|---|-----------------------------------------------------------------|
| 1 | Title cover — deck title, subtitle, author, date                |
| 2 | Agenda / TOC — numbered sections with hairline dividers         |
| 3 | Section break (big text) — eyebrow + large phrase + accent rule |
| 4 | Section break (big number) — giant numeral + section title      |
| 5 | Pull quote — large quote + attribution                          |
| 6 | Closing / thank-you — headline + call to action + contact line  |

Shares the same brand constants as the standard templates so the two
decks look like they belong to the same family.

## Build

```
pip install python-pptx Pillow
python3 examples/build_standard_templates.py     # → examples/standard_templates.pptx
python3 examples/build_title_and_breaks.py       # → examples/title_and_breaks.pptx
python3 authoring/cli.py ingest examples/standard_templates.pptx
python3 authoring/cli.py ingest examples/title_and_breaks.pptx
```

The standard-templates script builds real `python-pptx` chart / table
/ picture primitives (not stand-in rectangles), so when you
`cli.py ingest` the resulting deck, the right-side slots categorise
as `chart`, `image`, and `table` correctly.

## Re-brand

Edit the colour and font constants at the top of both build scripts
(keep them in sync so the two decks stay a family):

```python
ACCENT = RGBColor(0x1F, 0x4E, 0x8C)   # primary accent
INK    = RGBColor(0x22, 0x2A, 0x35)   # body text
MUTED  = RGBColor(0x6B, 0x73, 0x82)   # captions / footers
TITLE_FONT = "Calibri"
BODY_FONT  = "Calibri"
```

Re-run the scripts to regenerate the decks.

## Cleanup

Decks and their derived artifacts live under
`authoring/workspace/` (gitignored):

```
workspace/
├── decks/<stem>/         deck folder + theme.yaml + slides/slide_NN.pptx
├── themes/<stem>/        v5 theme (one-to-one with deck on ingest)
├── skeletons/<sk>/       v5 skeleton.yaml + preview.png + background.png
└── assets/               extracted atoms shared across decks
    ├── <sha>.png/.jpg    raster images (consumed by v5 image slots)
    ├── <sha>.xml         table / chart / smartart / freeform OOXML
    └── <sha>.yaml        sidecar metadata (headers, sample cells, etc.)
```

### Remove a deck

```bash
python3 authoring/cli.py remove-deck <stem> [--dry-run]
```

Drops the deck folder, its v5 theme, and every skeleton sourced from
it. **Assets stay** — they're deduplicated by content SHA and reusable
across decks; only the dead `sources:` refs in each asset yaml are
pruned. Quote stems with spaces: `remove-deck "deck with spaces"`.

### Manual asset cleanup

There's no central asset registry (asset listing is a glob of
`workspace/assets/*.yaml` at read time), so just `rm` what you don't
want:

```bash
cd authoring/workspace/assets
rm asset_<id>.png asset_<id>.yaml             # specific image
rm *.xml                                      # all XML atoms (see below)
```

### XML atom purge (safe today)

The `.xml` asset binaries — extracted table / chart / SmartArt /
freeform OOXML fragments — are **not consumed by current compose
paths**:

- v5 compose builds tables and charts from python-pptx primitives
  (`add_table` / `add_chart`); the XML is ignored.
- v4 compose deep-copies tables but explicitly skips chart and
  smartart atoms ("not yet placeable").

The sidecar `*.yaml` files are still useful for agent reasoning
(table headers, chart series names, etc.). Purge the binaries
without losing metadata:

```bash
cd authoring/workspace/assets
rm *.xml
```

…or drop both binary and sidecar if the metadata isn't worth
carrying either:

```bash
cd authoring/workspace/assets
for f in *.xml; do rm "${f}" "${f%.xml}.yaml"; done
```

Raster images (`*.png`, `*.jpg`) are real content — keep those.
