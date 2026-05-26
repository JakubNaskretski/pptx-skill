# Standard templates deck

Ten neutral-brand 16:9 layouts to feed the ingest pipeline, so the
agent has a known shape vocabulary to compose against. Each slide
has the same backbone: header at top, text or bullets on the left,
chart / image / table on the right.

## Layouts

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

## Build

```
pip install python-pptx Pillow
python3 examples/build_standard_templates.py
# → examples/standard_templates.pptx
python3 authoring/cli.py ingest examples/standard_templates.pptx
```

The script builds real `python-pptx` chart / table / picture
primitives (not stand-in rectangles), so when you `cli.py ingest`
the resulting deck, the right-side slots categorise as `chart`,
`image`, and `table` correctly.

## Re-brand

Edit the colour and font constants at the top of the build script:

```python
ACCENT = RGBColor(0x1F, 0x4E, 0x8C)   # primary accent
INK    = RGBColor(0x22, 0x2A, 0x35)   # body text
MUTED  = RGBColor(0x6B, 0x73, 0x82)   # captions / footers
TITLE_FONT = "Calibri"
BODY_FONT  = "Calibri"
```

Re-run the script to regenerate the deck.

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
