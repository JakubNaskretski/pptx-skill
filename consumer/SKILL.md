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

## Filterable metadata

Both templates and assets expose:

- `intent` (templates) / `subject` (assets) — what it is
- `feel` — `formal`, `punchy`, `data-dense`, `warm`, `clinical`,
  `minimal`, `playful`, `dramatic`, `celebratory`
- `suitable_for` — `opener`, `section_divider`, `content`, `data`,
  `quote`, `closing`, `product`, `team`, `hero`, `culture`, `event`,
  `abstract`, `decorative`
- `kind` (assets) — `photo`, `icon`, `logo`, `illustration`,
  `screenshot`
- `composition` (assets) — `centered`, `left-weighted`,
  `right-weighted`, `full-bleed`, `top-heavy`, `scattered`
- `colors` (assets) — dominant palette words

`--filter` accepts comma-separated `key=value` pairs. List values
match if any element equals the filter value.

## Files

```
SKILL.md              this file
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
