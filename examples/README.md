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
