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

### Mode 2 — compose from atoms (v4) — full freedom mode

Build a slide from scratch on a blank canvas. You can place **any
atom from the catalog** (photo, logo, illustration, vector, table,
callout, freeform, smartart) at any position, mixed with **native
textboxes** you write yourself. The host template's theme (colours +
fonts) still applies to the slide — see `theme_colors` / `fonts`
on the host you picked.

Use this mode whenever a template's `slots` don't expose what you
need. Want a photo where the template only has a text slot? Want
to mix a logo + table + freeform title that don't naturally
co-occur in any single template? Want to add a one-off text
annotation positioned freely? Compose-mode is the answer — it is
not a fallback, it is a peer to template-mode. Pick template-mode
when an existing template *already encodes the layout you want*;
pick compose-mode otherwise.

```json
{
  "compose": true,
  "layout": "title-top-2col",
  "shapes": [
    {"kind": "text", "value": "Pipeline view",
     "x": 0.05, "y": 0.05, "w": 0.9, "h": 0.1,
     "bold": true},
    {"atom": "asset_8c14", "kind": "image",
     "x": 0.05, "y": 0.2, "w": 0.4, "h": 0.5},
    {"atom": "asset_cb3a", "kind": "callout",
     "x": 0.05, "y": 0.75, "w": 0.4, "h": 0.15,
     "recolor": {"#ff0000": "accent"}},
    {"atom": "asset_4f01", "kind": "table",
     "x": 0.55, "y": 0.2, "w": 0.4, "h": 0.6,
     "cells": [["Quarter","Revenue"], ["Q3","$1.2M"], ["Q4","$1.8M"]]}
  ]
}
```

The shape spec is uniform regardless of asset kind — pass any
`asset_<id>` from the bundle's `assets` index along with fractional
geometry. Picture atoms route through `add_picture`; XML atoms
(table, callout, freeform) graft as fragments. The `kind` field on
each shape is informational + lets the engine apply kind-specific
handling like `cells` for tables.

### How to choose between modes

- **Template-mode** when an existing template `slots[]` + `layout`
  already match the slide you want — fill the slots, ship it. Best
  for "standard slide types" the deck author already pre-designed:
  openers, two-column layouts, dashboards, closers.
- **Compose-mode** when the layout you need doesn't exist as a
  template, OR when you want to mix atoms from multiple decks onto
  a single slide (e.g. a photo atom from deck A on a host template
  from deck B), OR when you want full positional control. Don't
  force a template to do something it wasn't designed for — drop
  into compose-mode instead.
- **Hybrid plan** — a plan is an array, so you can mix entry kinds
  freely. Most decks end up template-mode for 60-80% of slides and
  compose-mode for the rest where flexibility matters.

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
- `layout` — a richer spatial summary string of *every* visually-
  meaningful shape on the slide (placeholders, pictures, atoms),
  not just placeholders. Tokens are `label@region` where region is
  one of `top|middle|bottom`-`left|center|right`. Repeats of the
  same kind get suffixed (`image@…, image#2@…`). Example:
  `"title@top-center, image@right, bullets@left, image#2@bottom-right"`.
- `inventory` — per-shape dicts for the assets that appear on this
  slide. Each entry: `{atom, kind, x, y, w, h, region}` where x/y/w/h
  are slide-relative fractions (0.0-1.0). Lets you pick a template
  not just by intent but by *anatomy* — "this template has a chart
  bottom-right and a callout middle". Same atoms are individually
  addressable from compose-mode via their `atom` id.

Assets additionally expose kind-specific blocks (v4):

- `table`: `{rows, cols, headers, sample_cells}`
- `chart`: `{type, series_count, categories_count}`
- `shape`: `{geometry, is_recolorable}` (for callout / freeform)
- `smartart`: `{layout, nodes}`
- `recolor_targets` (vector): hex codes the engine can rewrite

## Informational fields (v4.1) — not filterable

Both templates and assets can carry an `interpretation` string when
the describing model surfaced soft observations alongside the strict
descriptive fields. Treat this as **context to consider, never as a
filter target or routing rule**:

- It carries hunches, plausible-but-uncertain identifications, era
  guesses, design rationale the model inferred, cross-asset patterns
  it noticed, etc.
- `--filter interpretation=…` is not supported and should not be
  attempted — the field is deliberately excluded from the strict
  retrieval pipeline so a wrong hunch cannot mis-route a pick.
- Use it the way a human designer would use marginalia: it can tip
  ties between otherwise comparable picks (e.g. a brief mentioning
  "early-20th-century physics" matches an asset whose
  `interpretation` reads "appears to be Einstein in 1920s Berlin",
  even though `subject` only says "older man at a chalkboard"), but
  never override the strict fields.

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
user_assets/          optional — present only in prompt-bundles when the
  <id>.<ext>          user attached their own images / SVGs / XML atoms
  manifest.json       to THIS compose request. See section below.
helpers/              optional — read-only kb_* Python utilities the
                      agent can run against the bundle (filter, inspect,
                      lint). Present in prompt-bundles, absent from
                      cli-build skill.zip.
reader.py             three-command entry point
requirements.txt      python-pptx, pyyaml
```

`index.json` is regenerated on every authoring `build`. Treat it as
the canonical filter index — individual sidecars carry the full
detail.

## User-supplied assets (prompt-bundle only)

When `user_assets/` is present in the bundle, the user has attached
their own assets to this specific compose request. They are a stronger
signal than KB catalog matches — the user expects these to appear in
the output.

- `user_assets/manifest.json` lists each entry with id, original
  filename, kind, and original dimensions.
- Files at `user_assets/<id>.<ext>` are LOW-RES previews. The user's
  machine holds the originals and splices them in at compose time —
  the agent doesn't need to (and cannot) load the full-res versions.
- Ids use the same `asset_<8>` format as the catalog. Reference them
  the same way: `"<slot>": "<id>"` for image slots, or
  `{"atom": "<id>", ...}` inside a compose-mode shape. The compose
  pipeline resolves both KB and user assets through the same path.
- User assets carry NO description. Read the brief, look at the
  preview, infer intent.
- If a user asset doesn't fit any slot: pick a different template,
  drop into compose-mode, or (last resort) substitute with the closest
  KB asset. Don't silently drop user assets.

## Backward compatibility

All v3 plans continue to work unchanged. The new shapes (slot value
polymorphism, compose-mode entries, theme awareness) are additive.
Until Phase D ships, the new shapes are accepted defensively but
their advanced features (colour overrides, recolouring, atom
assembly) are dropped with one-line warnings rather than rendered.
