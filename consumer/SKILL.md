---
name: pptx-skill
description: Compose brand-aligned PowerPoint decks from a curated library of pre-described slide templates and visual atoms. Pick templates by intent + feel, fill text and image slots, output a .pptx. No vision capability required at consumption time — every asset, template, and atom has a pre-written description.
---

# pptx-skill

Portable presentation-template library. Pick pre-described slide
templates and atoms by reading text metadata; compose a deck.

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
   front — palette / voice / style constraints. Apply to every text
   choice and template pick.
1. **Read deck themes (v4).** Each ingested deck ships a
   `decks/<deck>/theme.yaml` with its actual palette + fonts — useful
   when reasoning about style fit between templates.
2. **Plan the deck.** Decide slide sequence: opener → content → closer.
3. **Find templates.** `list --filter suitable_for=opener,feel=punchy`
   returns matching items.
4. **Find atoms.** Same for image / table / callout slots:
   `list --filter kind=photo,suitable_for=team`. New v4 kinds:
   `vector`, `table`, `chart`, `callout`, `freeform`, `smartart`.
5. **Build a plan.** JSON array of entries — see "Plan format" below.
6. **Compose.** `compose plan.json out.pptx`.

## Plan format

A plan is a JSON **array** of entries. Each entry is one of two kinds.

### Mode 1 — template + slots (existing, recommended for most slides)

Pick a pre-described template, fill its slots.

```json
{
  "template": "deckQ3_03",
  "slots": {
    "title": "Q4 results",
    "hero": "asset_b91f22"
  }
}
```

Slot keys must match the template's declared `slots` (call
`get <template_id>` to inspect them).

### Mode 2 — compose from atoms (v4)

Assemble a slide from atoms on a blank canvas. Use when no template
fits but you want to reuse captured atoms.

```json
{
  "compose": true,
  "layout": "title-top-2col",
  "shapes": [
    {"kind": "text", "value": "Pipeline view",
     "x": 0.05, "y": 0.05, "w": 0.9, "h": 0.1,
     "bold": true},
    {"atom": "asset_cb3a", "kind": "callout",
     "x": 0.05, "y": 0.2, "w": 0.4, "h": 0.4,
     "recolor": {"#ff0000": "accent"}},
    {"atom": "asset_4f01", "kind": "table",
     "x": 0.55, "y": 0.2, "w": 0.4, "h": 0.6,
     "cells": [["Quarter","Revenue"], ["Q3","$1.2M"], ["Q4","$1.8M"]]}
  ]
}
```

**Status (current build):**
- `kind: text` shapes render as plain textboxes; `bold` is honored.
  `font_role` / `color_role` are accepted but warn (not yet honored).
- Atom placements with `x/y/w/h` (fractions of slide) work for
  pictures, vectors, tables, callouts, and freeforms. `recolor`
  rewrites `<a:srgbClr>` fills in the atom; role tokens (`accent`,
  `primary`, `text`, `background`) resolve against the host
  template's theme colors.
- Tables accept an optional `cells` override (same semantics as a
  `kind: table` slot).
- Chart and SmartArt atoms are **skipped with a warning** —
  related-parts copying is deferred. Use the picture export of those
  atoms instead (if you have one).
- The host deck is the first template-mode entry's pptx (so its
  master / theme apply). If the plan is all compose-mode entries,
  the first template in the bundle (alphabetical) becomes scratch
  host and its original slide is dropped.
- **Cross-deck normalisation (D5).** When copying a template or atom
  from a non-host deck, semantically-named scheme slots (`primary`,
  `accent`, `text`, `background`) get remapped from the source deck's
  alias to the host deck's alias so a brand colour stays a brand
  colour. Aspect-ratio mismatches between source and host warn once
  per source deck but don't refuse or auto-scale.

## Slot value polymorphism (v4)

Slot values can take any of these shapes:

| Shape | Use | Honored today? |
|---|---|---|
| `"text content"` | string — plain text | ✅ |
| `["a", "b"]` | bullets — array of plain strings | ✅ |
| `"asset_<id>"` | image — references an asset by id | ✅ |
| `[["h1","h2"], ["r1c1","r1c2"]]` | table — list-of-lists of cell strings (for `kind: table` slots) | ✅ |
| `{"cells": [[...], ...]}` | table — same data wrapped (other keys ignored with a warning) | ✅ |
| `{"text": "...", "color_role": "accent", "bold": true}` | styled text | ⚠️ degraded to plain text |
| `{"runs": [{"text": "X", "bold": true}, {"text": " Y"}]}` | per-run rich text | ⚠️ degraded to concatenated text |
| `{"asset": "asset_<id>", "recolor": {"#ff0000": "accent"}}` | image with overrides | ⚠️ degraded to plain image |

⚠️ = accepted without crashing; styling dropped with a one-line
warning per slot. Phase D will honor these fully. **Existing v3
plans keep working unchanged.**

`color_role` values: `primary`, `accent`, `text`, `background`, or any
clrScheme slot from theme.yaml (`accent1`, `dk2`, `lt1`, …).

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
- `kind` (assets) — pictures: `photo`, `icon`, `logo`,
  `illustration`, `screenshot`. Vector: `vector`. Structured atoms:
  `table`, `chart`, `callout`, `freeform`, `smartart`.
- `composition` (assets, pictures only) — `centered`, `left-weighted`,
  `right-weighted`, `full-bleed`, `top-heavy`, `scattered`
- `colors` (assets) — dominant palette words
- `colors_hex` (assets, v4) — actual hex codes from binary inspection

Templates additionally expose (v4):

- `theme_colors` — `{primary, accent, text, background}` resolved
  from the source deck's theme via aliases
- `fonts` — `{major, minor}` from the deck's theme font scheme
- `slots[].style` — captured run-level overrides (`size_pt`, `bold`,
  `color_role`, `font`, …) where the template explicitly set them
- `inventory` — list of atom ids carried by this template (so you
  know what comes "for free" when picking it; the same atoms are
  individually addressable from compose-mode)

Assets additionally expose kind-specific blocks (v4):

- `table`: `{rows, cols, headers, sample_cells}`
- `chart`: `{type, series_count, categories_count}`
- `shape`: `{geometry, is_recolorable}` (for callout / freeform)
- `smartart`: `{layout, nodes}`
- `recolor_targets` (vector): hex codes the engine can rewrite

`--filter` accepts comma-separated `key=value` pairs. List-valued
fields match if any element equals the filter value. Within a
single key, `|` expresses OR: `feel=warm|formal` matches either.
The literal value `none` matches items where the field is missing
or empty — use it to include un-tagged items (`feel=warm|none`)
or to find them explicitly (`feel=none`).

## Bullets — do not prepend bullet glyphs

PowerPoint templates apply bullets via layout formatting. Do **not**
prepend `•`, `-`, or `*` to lines — you'll get double bullets in the
rendered slide.

- `kind: bullets` slot — array of plain strings. The template adds the
  bullet character.
- `kind: text` slot that displays as a list — multi-line string joined
  by `\n`, no leading glyph.

`compose` defensively strips leading bullet glyphs, but correct input
keeps your plan readable.

## Per-deck theme (v4)

Each deck ships its theme metadata at `decks/<deck>/theme.yaml`:

```yaml
palette:
  dk1: "#000000"
  lt1: "#FFFFFF"
  accent1: "#FF388C"
  …
aliases:
  primary: dk2
  accent: accent1
  text: dk1
  background: lt1
fonts:
  major: Helvetica
  minor: Century Gothic
aspect: "16:9"
```

Use this to reason about which templates fit a brand policy. The
`aliases` map gives you friendly names for any clrScheme slot.

## Files

```
SKILL.md              this file
brand.md              optional — palette/voice/style rules to apply when composing
index.json            flat list of all templates + assets + tags
decks/<deck>/
  theme.yaml          v4 — per-deck palette + fonts + aspect (informational)
templates/<id>/
  slide.pptx          reusable slide fragment
  meta.yaml           full sidecar (intent, slots, suitable_for, …)
  preview.png         optional
assets/<id>.<ext>     binary — png/jpg/svg/emf/xml (xml = structured atom)
assets/<id>.yaml      full sidecar
reader.py             three-command entry point
requirements.txt      python-pptx, pyyaml
```

`index.json` is regenerated on every authoring `build`. Treat it as
the canonical filter index — individual sidecars carry the full
detail.

## Backward compatibility

All v3 plans continue to work unchanged. The new shapes (slot value
polymorphism, compose-mode entries, theme awareness) are additive.
Until Phase D ships, the new shapes are accepted defensively but
their advanced features (colour overrides, recolouring, atom
assembly) are dropped with one-line warnings rather than rendered.
