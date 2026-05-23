# pptx-skill — deferred follow-ups

Items decided to be worth doing but explicitly deferred during the
FINDINGS.md walkthrough. Each entry links back to the finding it
came from, names the options on the table, and notes the call we
made on timing.

---

## A1.3 — Mixed-master compose: host selection + theme normalization

**Source:** FINDINGS.md → Part A → A1.3 (first-template-as-host
compose strategy, `consumer/reader.py` `cmd_compose` ~lines 410-447).

**Problem:** the first template in the plan silently becomes the
"host" — its master slide, theme colors, fonts, and aspect ratio
are applied to every subsequent slide that's deep-copied in. Mixing
templates from decks with different masters/themes/aspect ratios
produces font/color drift and (for aspect mismatches) cropping or
distortion, with no warning to the user.

### Option 1 — Lightweight: explicit host + compatibility warnings

Small, defensive, no XML rewriting.

- `reader.py compose` gains `--host <template_id>` to override the
  default of "first in plan".
- At compose start, gather aspect ratio + theme/master identifier
  of every referenced template. If any disagree with the host,
  emit a stderr warning naming the offenders.
- Optional `--strict` flag to fail instead of warn.

**Value:** the user sees the problem instead of getting a quietly
broken deck. Doesn't fix anything, but makes the failure mode
visible and gives an escape hatch.

### Option 2 — Full theme normalization (the real fix)

When deep-copying a foreign slide's shape tree into the host:

- Rewrite theme-color references (`<a:schemeClr val="accent1"/>`
  etc.) so they map from the foreign theme's slot semantics into
  the host theme's. Requires reading both `theme.xml` files and
  building a mapping (probably by role name + perceptual nearest
  color when role names don't align).
- Same treatment for theme fonts (`+mj-lt`, `+mn-lt` → host's
  major/minor latin fonts).
- For aspect mismatches: either refuse to compose, or scale +
  reposition shapes proportionally into the host's slide
  rectangle. Decide policy.
- Master-level decoration (page numbers, brand bars) from the
  host's master will already apply automatically once shapes are
  on a host-derived layout — but see A1.4, which today strips
  layout placeholders before the copy. That needs fixing in
  concert (don't strip; merge).

**Value:** templates from any deck compose into any host without
visual breakage. Unlocks the "library across multiple brand
decks" use case, which is currently aspirational.

**Cost:** real XML work. Probably 1-2 focused days. Tests are
mandatory — this is exactly the kind of change that silently
breaks decks if untested. Tie to FINDINGS A1.6 (no tests).

### Also in scope: A1.4 — stop stripping layout placeholders

**Source:** FINDINGS.md → A1.4 (`consumer/reader.py:299-301`).

Today `_copy_slide_into` clears every shape on the chosen layout
before pasting the foreign slide's shapes in. That wipes
master-level decoration (page numbers, brand bars, footers) from
every appended slide. Plus the layout-picking heuristic at
`reader.py:286-295` ("pick the layout with the fewest
placeholders") is arbitrary and non-deterministic across decks.

Belongs with the normalization workstream: the whole point of
normalizing onto the host is to keep the host's master decoration
visible. The fix is to **merge** foreign shapes onto a
host-derived layout instead of clearing-then-pasting, and to pick
the target layout deliberately (e.g. by intent/role, or let the
plan specify it) rather than by shape-count heuristic.

### Decision

**Lightweight first, full normalization after the walkthrough.**
The user has explicitly flagged full normalization as important and
wants to address it as a focused workstream once the FINDINGS.md
pass is complete. Lightweight version is a sensible interim — even
post-normalization, `--host` is still useful as a UX affordance.
A1.4 ships as part of the normalization workstream — fixing
host-master inheritance is meaningless if we keep wiping the
layout it provides.

---

## Non-text / non-picture element handling

**Sources:** FINDINGS.md → A2.7 (`PP_PLACEHOLDER.OBJECT` treated as
textual, `authoring/cli.py:113-119`); Part B → B1 (inventory of shape
kinds we ignore); Part B → B2 (SVG fallback).

**Problem:** today the skill recognises exactly two asset-bearing
shape kinds — textual placeholders and raster `Picture` shapes
(>20% slide area). Everything else either rides along invisibly
inside the slide XML (auto-shapes, freeforms, groups, lines,
charts, tables, SmartArt) or is silently downgraded (SVG → raster
fallback). The index can't filter on them, the plan can't address
them, compose can't swap or recolour them.

### Specific gaps

- **`OBJECT` placeholders mis-typed as text.** Currently bundled
  with TITLE/BODY in `PLACEHOLDER_TEXTUAL`. Empty OBJECT
  placeholders happen to behave fine (they do hold text), but
  filled OBJECT placeholders (chart/table) will crash ingest —
  `shape.text_frame` raises in python-pptx and the current
  `getattr(..., None)` doesn't catch it.
- **Charts.** Not captured as assets. Carried through compose
  via deep-copy of slide XML but unaddressable.
- **Tables.** Same as charts. No way for a plan to swap cell
  values or even reference the table as a slot.
- **SmartArt.** Partial XML survival; not extracted as an asset;
  no way to retarget.
- **Auto-shapes / freeforms / callouts / connectors / groups /
  lines.** Ignored by `extract_picture_assets`. A deck full of
  callout arrows or annotated shapes loses all of that structure
  to the index — agents can't pick a slide *because* it has the
  right callout pattern.
- **SVG vector assets.** PPTX stores SVG `Picture` shapes with
  both the SVG and a raster fallback. python-pptx's `shape.image`
  returns the fallback; today's `extract_picture_assets` writes
  the PNG/EMF and discards the SVG. Visual result is OK; vector
  source is gone, blocking recolour / DPI-independent re-render.
- **Theme/master backgrounds.** Ride along via master; no
  separate signal to the agent.

### Options on the table

- **Minimum defensive (small):** wrap `text_frame` access in
  `try/except` so filled OBJECT placeholders don't crash ingest;
  log and skip. ~3 lines. Doesn't add any new capability — just
  removes a sharp edge.
- **Classify only (medium):** detect `has_chart` / `has_table` /
  `has_smart_art` at ingest and emit dedicated `kind: chart |
  table | smartart` slots. Compose still can't fill them — but
  the index becomes filterable ("slides with charts"), and we
  stop pretending they're text.
- **Capture as assets (large):** extend `extract_picture_assets`
  into a general `extract_assets` that recognises non-picture
  shape kinds. Decide for each kind what the asset binary even
  *is* — for SVG, dig out the `asvg:svgBlip` rel; for shapes,
  probably serialise the shape XML as the asset payload. New
  asset `kind` values, schema work, scope-creep risk.
- **Full round-trip (largest):** plan format learns to *swap* a
  chart's data, a table's cells, a SmartArt node's text. This
  is well past the current architecture; would need a separate
  design pass.

### Decision

Deferred. Tightly coupled with the colour / structured-palette
work in Part B (B3-B5) — both touch ingest and both want a more
structured slot/asset model. Worth deciding the schema shape
(plan slot value as string-or-object, asset `kind` extensions)
*once*, then doing classify-and-capture in one pass rather than
shipping defensive patches now and rewriting them later.

Minimum-defensive 3-line `try/except` around `text_frame` may be
worth doing inline if a real deck hits the crash — currently
theoretical, no reported failure.

---

## A2.10 — Smarter slot-vs-decoration classifier for free pictures

**Source:** FINDINGS.md → A2.10 (`authoring/cli.py:228`, hardcoded
20% slide-area threshold).

**Problem:** during ingest, free `Picture` shapes (non-placeholder
raster images) are classified as "slot" if they cover >20% of the
slide area, otherwise "frozen decoration". The boundary is hidden,
arbitrary, and has no escape hatch:

- A 19%-area photo silently becomes background — agent can never
  swap it.
- A 21%-area logo silently becomes a swappable `hero` slot.
- No CLI flag, no per-shape override, no signal in the resulting
  YAML that classification was even attempted.
- Decks where most photos happen to land near the threshold
  produce empty `slots: []` with no diagnostic.

### Multi-signal design

Replace the single area threshold with a small ranked classifier
combining:

- **Area** (current signal). Above ~25% → strong slot signal;
  below ~5% → strong decoration signal; mid-band needs other
  signals.
- **Position.** Corners and edges (top-N%, bottom-N%, left/right
  margin bands) bias toward decoration. Mid-slide biases toward
  slot.
- **Aspect.** Extreme aspects (very wide banners, very thin
  strips) bias toward decoration. Slot-like aspects (16:9, 4:3,
  1:1, 3:4, 9:16) bias toward slot.
- **Repeat-across-slides.** Pictures whose SHA1 appears on >50%
  of slides are almost certainly brand decoration (logo, header
  mark, page-corner badge). Strong decoration signal regardless
  of size.

Combine into a confidence score; emit slot if above threshold.
Write the classification (and the signals that drove it) into
the slide.yaml as metadata so a human can audit "why was this
treated as decoration".

### Architectural change required

Today `detect_slots` is per-slide and stateless. The
repeat-across-slides signal needs cross-slide knowledge. Two
shapes:

- **Pre-pass.** Walk all slides once to build
  `{sha1 → slide_count}`; pass into `detect_slots` as a
  `repeat_map`.
- **Two-phase ingest.** Collect raw shape inventory, then
  classify in a second pass. Cleaner separation; bigger
  refactor.

Pre-pass is the smaller change.

### Escape hatches (regardless of classifier complexity)

- CLI flag `--slot-min-area <fraction>` to override the threshold
  for the whole deck. Cheap interim if the classifier is delayed.
- Per-shape override via shape name prefix (e.g. `slot_` to force
  inclusion, `frozen_` to force exclusion) so users can correct
  the classifier without editing YAML by hand. Requires
  documenting the convention.

### Tests are mandatory

The whole point of this rework is *better* classification. Without
fixtures (a hand-labelled deck with known expected slots), any
new heuristic is just differently arbitrary. Tie to A1.6 — this
is one of the workstreams that justifies bringing tests up.

### Decision

Deferred. Bundles naturally with the
[[Non-text / non-picture element handling]] workstream and the
Part B palette / structured-formatting work — all three touch
`detect_slots` / `extract_picture_assets` and want a richer
ingest model. One coordinated rework with tests beats three
serial patches.

No interim `--slot-min-area` flag landed; revisit if a real deck
hits the threshold pathologically.

---

## Workspace lifecycle: removal of KB items

**Source:** FINDINGS.md → A2.13 (no way to remove a slide from the
workspace).

**Problem:** workspace state is effectively append-only. Re-ingest
is idempotent for content that still exists in the source deck,
but nothing in the toolchain ever *removes* a knowledge-base item.
Cases that go wrong silently:

- **Slide deleted from source deck.** The old `slide_NN.yaml`
  lingers — agent still sees it as a valid template. Worse: the
  fragment `.pptx` may also linger, and if numbering shifted in
  the source, the slot positions in the orphan YAML now refer to
  a different fragment.
- **Asset replaced in source deck.** New SHA1 → new asset file is
  written; old SHA1-named asset stays on disk and in
  `asset_index.yaml`, still pickable by the agent.
- **Source deck removed entirely.** All its slides + assets
  linger. No `sources:` reconciliation.
- **Human-rejected slide.** A user reviewing the workspace might
  decide "this template is bad, never pick it" — there's no
  affordance for that except deleting the YAML by hand, which
  ingest will silently regenerate next run.

### What's needed

A coherent **lifecycle model** for KB items, not just one-off
delete commands. Rough shape:

- **Re-ingest as reconciliation, not append.** Compare current
  source-deck slide set against existing workspace YAMLs; mark
  orphans (item exists in workspace but not in any current
  source). Either auto-prune with a confirmation, or move to a
  quarantine dir (`_removed/`) the agent doesn't read.
- **Explicit removal commands.** `cli.py remove-slide <id>`,
  `cli.py remove-asset <id>`, `cli.py remove-source <deck>`.
  Each cleans up YAMLs, fragments, asset binaries, and entries
  in `asset_index.yaml` atomically. Idempotent.
- **Soft-reject affordance.** A `rejected: true` (or
  `status: rejected`) field on slide/asset YAML that excludes
  the item from agent retrieval without deleting it. Survives
  re-ingest. Lets humans curate without losing the data.
- **Garbage collection.** Sweep that finds unreferenced asset
  binaries on disk (no slide YAML mentions them, no
  `asset_index.yaml` entry) and removes them.

### Coupling

- **Idempotence guarantees.** Reconciliation must not destroy
  hand-edited descriptive fields on items that *do* still exist.
  Diff at the structural level (id, source deck, source slide
  index), preserve descriptive fields (`feel`, `intent`,
  `colors`, `scope`, etc.) verbatim.
- **Multi-source decks.** A slide's identity is
  `(deck, slide_index)`. Two decks with overlapping content
  shouldn't get reconciled into each other.
- **Asset reference counting.** Removing a slide that referenced
  `asset_abc12345` shouldn't blindly delete the asset — other
  slides might still reference it. Need ref-count or
  reachability walk (mirrors the fragment-GC pattern from
  A1.5/A2.12).
- **Tests are mandatory.** This is exactly the
  "silently-destroys-user-work" class of change that
  motivates A1.6.

### Decision

Deferred. Real workstream — touches CLI surface, ingest flow,
schema (rejected field), and needs the same testing infra as
the normalization workstream. Worth designing the lifecycle
model in one pass rather than bolting on `remove-slide` and
discovering ref-counting + reconciliation problems halfway
through.

Interim workaround for users today: delete the YAML + fragment
by hand; re-ingest will *not* regenerate items whose source
slide was also deleted from the deck. That covers the most
common case ugly-but-functionally.

---

## Colour / formatting / SVG: the structured-style workstream

**Sources:** FINDINGS.md → A3.15 (`font.color` not copied at compose,
`consumer/reader.py:128-200`); Part B → B2 (SVG fallback only);
B3 (plan / schema / compose all colour-blind);
B4 (no rich runs / per-word formatting); B5 (no two-way YAML).

**Problem:** today the toolchain is fundamentally colour-and-style
agnostic end-to-end. A deck with a brand-coloured title, a
mixed-formatting subtitle, an SVG logo, or any structured palette
goes in → comes out neutered.

### Specific gaps

- **Compose strips `font.color`.** `_fill_text_shape` and
  `_fill_bullets_shape` copy size/bold/italic/name but explicitly
  not colour. Brand-coloured titles fall back to default black /
  theme-default after any text swap. Single most visible
  regression in compose.
- **Plan can't express colour.** `_apply_slot_value` accepts only
  bare string, list of strings, or `asset_<id>`. No
  `{"color": "#..."}` branch, no theme-token branch
  (`{"color_role": "accent1"}`).
- **Plan can't express per-run formatting.** No way for an agent
  to say "bold the word `risk`" or "right-align this subtitle".
  Single style per slot, derived from the template's first run.
- **Schema has no structured palette.** `asset.yaml` carries
  colour *names* (`navy`, `warm white`) for retrieval, not hex.
  `slide.yaml` has no colour field. `brand.md` is free-text
  prose. Nothing to match against, nothing to enforce.
- **SVG is silently downgraded to raster.** PPTX stores SVG
  `Picture` shapes with both the SVG and a raster fallback;
  python-pptx's `shape.image` returns the fallback. Today's
  ingest writes the PNG/EMF, drops the SVG. Result *looks* OK
  but vector source is gone — no recolour, no DPI-independent
  re-render, no theme-aware logo swap.
- **No ingest-side colour extraction.** Even if compose could
  carry colour, nothing is extracting it. Theme palette
  (`master.element.xpath('.//a:clrScheme')`), per-run
  `font.color.rgb` / `font.color.theme_color`, asset dominant
  colours — all available from python-pptx + PIL but unused.
- **Compose nukes picture-shape styling.** `_replace_image_shape`
  (`consumer/reader.py:202-219`) removes the original Picture
  shape and re-adds via `add_picture`, preserving only geometry.
  Lost on every image swap: crop, picture border, rotation,
  shadow, transparency, picture effects (recolour / artistic),
  alt text. Brand decks with styled photo frames get stripped
  every compose. Direct analogue of A3.15 for images — same
  bug class (compose discards template fidelity), same fix
  shape (mutate the existing shape's image part instead of
  re-creating; let the plan override).
- **Bullet sub-runs lose template font + use the wrong
  line-break mechanism.** `_fill_bullets_shape`
  (`consumer/reader.py:197-199`) appends extra lines within a
  bullet as `"\n" + extra` sub-runs that skip the font-copy
  block entirely — so a multi-line bullet's line 1 has the
  template font but line 2+ falls back to paragraph defaults
  (visible as size/weight/name drift mid-bullet). Separately,
  stuffing `\n` into `run.text` is not the right way to express
  a soft line break in PPTX — the proper form is an `<a:br/>`
  element between runs. Both issues vanish once the structured
  workstream factors font-copy into a helper and the slot
  format learns about per-run / per-line style.

### Design shape (coordinated, not piecemeal)

Three coupled changes, worth deciding *together* before touching
code:

1. **Schema additions.** Structured palette in brand (YAML, not
   prose) and per-deck (`theme.yaml` or inlined into slide.yaml).
   Optional `colors: {hex, role}` on assets in addition to names.
   `theme_colors` on slide.yaml extracted at ingest from
   `master.color_theme`.
2. **Slot value polymorphism.** Today: `string | list[string] |
   "asset_<id>"`. Tomorrow: also accept `{"text": "...",
   "color": "#...", "color_role": "accent1", "bold": true,
   "runs": [...]}`. Reader needs a small dispatcher. Existing
   string-valued plans keep working — purely additive.
3. **Compose preserves template fidelity by default.** Two
   parts:
   - Text: in `_fill_text_shape` / `_fill_bullets_shape`, copy
     `font.color` from the template run (handle
     `MSO_COLOR_TYPE.RGB` and `SCHEME`); let the plan override
     per-run.
   - Images: rewrite `_replace_image_shape` to mutate the
     existing Picture shape's image part (swap the blob behind
     the same `<p:pic>`) instead of removing-then-re-adding.
     Keeps crop, border, rotation, shadow, transparency,
     effects, alt text. Plan can still override (e.g. a future
     `{"image": "asset_...", "alt": "..."}` shape).

### SVG (separate but adjacent)

`extract_picture_assets` digs into shape rels to pull both the
SVG and the raster fallback — write both, prefer SVG at compose
time, fall back to raster if the consumer can't handle SVG.
Requires `shape._element.xpath(...)` against `asvg:svgBlip`.
Belongs in the same workstream because it touches the same
ingest + asset-model surface.

### Why deferred together

Each of these *could* be done in isolation:

- The compose-colour fix is ~10 lines per function. Tempting to
  do inline.
- A `--copy-colors` flag could ship today as opt-in.

But shipping the compose-colour fix without the schema /
plan-format work means colour starts working *implicitly* (always
inherits from template) without giving the agent any way to
*choose* a colour. Then when the agent eventually wants to say
"use the accent for this title", we have to re-do compose to
handle the override case. Designing the polymorphic slot value
*first* means one coherent pass: compose learns to copy AND
override in the same change.

A1.6 (tests) is mandatory for this workstream — colour
regressions are exactly the silently-broken-deck class that
needs fixtures.

### Decision

Deferred. Real workstream — schema + plan format + compose +
ingest, all in one design pass. Tightly coupled with
[[Non-text / non-picture element handling]] (same ingest /
asset-model surface) and with FINDINGS.md "Suggested order of
attack" steps 1-3 (colour fix → palette extraction → plan-format
extension). Worth queuing them as one bundle.

If a user complains tomorrow about black titles in composed
decks, do the ~10-line compose-time copy as a stopgap and
document the limitation; the schema/plan work still happens.

---

## Done during the walkthrough

Items that were originally going to be deferred but ended up
getting done inline. Listed here so the normalization workstream
knows what prerequisites are already met.

### A1.5 / A2.12 — Fragment GC: full refactor (done 2026-05-23)

**Sources:** FINDINGS A1.5 (bare `except` around `drop_rel`) and
A2.12 (wasteful slide fragments — each `slide_NN.pptx` carried
the whole source deck's masters/layouts/media as orphan parts).

**What landed in `authoring/cli.py`:**

- `write_slide_fragment` now runs `prune_unreachable_parts`
  after `prs.save()`. The bare `except` around `drop_rel` is
  narrowed to `except KeyError`.
- `prune_unreachable_parts` does three things:
  1. For each slide master, drops `slideLayout` rel entries
     pointing to layouts no surviving slide actually uses
     (`_layouts_used_by_slides` + `_prune_master_rels`).
  2. Repairs each master XML by removing `<p:sldLayoutId>`
     entries whose `r:id` is no longer in the master's rels
     (`_repair_master_xml`). This handles both our active prune
     AND any orphan entries python-pptx left behind during its
     own save-time layout trimming. **Key gotcha:** `r:` inside
     document XML binds to `…/officeDocument/2006/relationships`,
     not the package `…/package/2006/relationships` used inside
     `.rels` files. Wrong namespace = drops every entry.
  3. Walks `/_rels/.rels` through every part's rels chain,
     collecting reachable parts; filters `[Content_Types].xml`
     Override entries to match; rewrites the zip keeping only
     reachable parts.

**Why this matters for the normalization workstream:** fragments
now carry exactly **one** master + **one** layout + **one** theme
+ only the media that slide actually uses. That removes the
ambiguity of "which theme does this template carry?" when the
normalizer needs to map foreign theme → host theme. Without this
cleanup, fragments could carry several masters/themes and the
mapping target wasn't well-defined.

**Measured impact (on a representative deck):** workspace 6488 KB
→ 6132 KB (~5.5%). Per-fragment the layout count dropped from 11
to 1. Bigger savings on heavier brand decks where unused layouts
carry meaningful media. Compose flow re-validated end-to-end.
