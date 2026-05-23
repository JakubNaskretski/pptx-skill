---
name: pptx-skill
description: Compose brand-aligned PowerPoint decks from a curated library of pre-described slide templates and visual assets. Pick templates by intent + feel, fill text and image slots, output a .pptx. No vision capability required at consumption time — every asset and slide has a pre-written description.
---

# pptx-skill

Portable presentation-template library. Pick pre-described slide
templates and assets by reading text metadata; compose a deck.

## Three commands

```bash
python reader.py list [--filter key=value,key=value]
python reader.py get <id>
python reader.py compose <plan.json> <output.pptx>
```

All commands write JSON to stdout. `compose` also writes the output
file.

## Agent workflow

0. **Read brand rules.** If `brand.md` is present, read it once up
   front. It carries palette / voice / style constraints that should
   shape every text choice and template pick.
1. **Plan the deck.** Decide slide sequence: opener → content → closer.
2. **Find templates.** `list --filter suitable_for=opener,feel=punchy`
   returns matching items.
3. **Find assets.** Same for image slots: `list --filter
   kind=photo,suitable_for=team`.
4. **Build a plan.** JSON array of `{template, slots}` entries.
5. **Compose.** `compose plan.json out.pptx`.

## Plan format

```json
[
  {
    "template": "deckQ3_03",
    "slots": {
      "title": "Q4 results",
      "hero": "asset_b91f22"
    }
  },
  {
    "template": "deckQ3_07",
    "slots": {
      "title": "Top three risks",
      "body": [
        "Risk one\nShort detail",
        "Risk two\nShort detail",
        "Risk three\nShort detail"
      ]
    }
  }
]
```

Slot keys must match the template's declared `slots` (call
`get <template_id>` to inspect them). Image-slot values are asset ids;
the reader resolves the binary path internally.

### Bullets — do not prepend bullet glyphs

PowerPoint templates apply bullets via layout formatting. Do **not**
prepend `•`, `-`, or `*` to lines — you'll get double bullets in the
rendered slide.

- `kind: bullets` slot — array of plain strings. The template adds the
  bullet character.
- `kind: text` slot that displays as a list — multi-line string joined
  by `\n`, no leading glyph.

`compose` defensively strips leading bullet glyphs from each line, but
correct input keeps your plan readable and avoids surprises.

## Filterable metadata

Both templates and assets expose:

- `intent` (templates) / `subject` (assets) — what it is
- `feel` — slide-feel and asset-feel are **separate** sets:
  - templates: `formal`, `punchy`, `data-dense`, `warm`, `clinical`,
    `celebratory`
  - assets: `formal`, `warm`, `clinical`, `punchy`, `playful`,
    `minimal`, `dramatic`
- `suitable_for` — also split by kind:
  - templates: `opener`, `section_divider`, `content`, `data`,
    `quote`, `closing`, `product`, `team`
  - assets: `team`, `hero`, `product`, `data`, `culture`, `event`,
    `abstract`, `decorative`, `closing`, `quote`
- `kind` (assets) — `photo`, `icon`, `logo`, `illustration`,
  `screenshot`
- `composition` (assets) — `centered`, `left-weighted`,
  `right-weighted`, `full-bleed`, `top-heavy`, `scattered`
- `colors` (assets) — dominant palette words

`--filter` accepts comma-separated `key=value` pairs. List-valued
fields match if any element equals the filter value. Within a
single key, `|` expresses OR: `feel=warm|formal` matches either.
The literal value `none` matches items where the field is missing
or empty — use it to include un-tagged items (`feel=warm|none`)
or to find them explicitly (`feel=none`). By default, missing or
empty fields do *not* match an explicit value.

## Files

```
SKILL.md              this file
brand.md              optional — palette/voice/style rules to apply when composing
index.json            flat list of all templates + assets + tags
templates/<id>/
  slide.pptx          reusable slide fragment
  meta.yaml           full sidecar (intent, slots, suitable_for, …)
  preview.png         optional
assets/<id>.<ext>
assets/<id>.yaml      full sidecar
reader.py             three-command entry point
requirements.txt      python-pptx, pyyaml
```

`index.json` is regenerated on every authoring `build`. Treat it as
the canonical filter index — individual sidecars carry the full
detail.
