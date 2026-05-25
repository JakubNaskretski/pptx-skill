# pptx-skill — v5 redesign: structural skeletons

Branch: `redesign/structural-slides` (off `main`). v5 is a
directional pivot, not an incremental v4.x release. v3 and v4 stay
on `main` until v5 is end-to-end working.

---

## tl;dr

v3 + v4 treat a slide as a *visually-faithful template*: we
deep-copy its XML into the output deck and the agent only swaps
text/image into slots. The describe step has been the load-bearing
piece — and a fragile one — because slides have to be semantically
described by a vision LLM.

v5 changes the unit. A slide becomes a **structural skeleton + functional
category**. The vision LLM stops describing slides; it only describes
individual asset binaries (icons, photos). Slides are categorized by
*function* (opening / data / closing / …) via a Flask UI step. On
approval we **digest** the slide: strip the styled foreground shapes,
keep a typed slot inventory with geometry, style hints, and
constraints. The agent gets the asset descriptions + the skeleton
catalog + a small set of new content-first selection methods, and
emits a JSON plan as today. The build engine constructs slides from
primitives on a chosen **host theme** instead of deep-copying source
XML.

## The trade

| Give up | Gain |
|---|---|
| Visual fidelity to source decks (specific layouts, custom callouts, designer polish) | Real constraint contract with the agent (max_chars enforced, not hinted) |
| The atom catalog for callouts / freeforms / SmartArt | A filterable catalog where "give me a 2-col comparison data slide" actually works |
| Vision-LLM slide description (the fragile load-bearing piece) | Workflow that resumes well: categorize in minutes, describe assets at your pace |
| `_copy_slide_into`, `_place_atom` deep-copy machinery | No cross-deck theme / font / aspect normalization to maintain (v4.1 items 1 + 3 become moot) |

Visual fidelity is not zero. Deck-level **theme** (palette, fonts,
master with brand bars / page numbers / corner logo / section panels)
is captured per deck and applied at build time. The agent picks **one
host theme** per output deck; skeletons are theme-free.

---

## Conscious design choices (do NOT relitigate)

Decisions locked during the v5 design conversation. Each has rationale
below; revisit only with reason. These are written down here
specifically so the next person reading the doc doesn't think they were
forgotten.

### Drop SmartArt, callouts, freeforms, auto-shapes from the atom catalog

**Conscious choice.** v4 captured these as atoms in case the agent
wanted to reuse them. In the v5 model the agent picks by *function and
constraint* (skeleton category, slot fit), not by visual look. These
atom kinds don't reduce to a useful constraint contract:

- **SmartArt** is a layout engine — PowerPoint owns the rendering,
  the data part defines nodes, the layout part renders them.
  Synthesizing equivalent SmartArt from scratch is fragile. If the
  user needs a process flow, the agent describes intent and the build
  engine assembles rectangles + arrows from primitives. Lossy but
  consistent.
- **Callouts / freeforms / auto-shapes** are visual fidelity. "Use
  this exact arrow shape" is the template-fidelity model we're
  explicitly leaving.

### Keep tables and charts — they ARE skeletons

Tables: row/col counts, header presence, per-cell content. Charts:
chart type, max series, max categories. Both are fully auto-ingestable
via python-pptx, both have a clean constraint contract, both let the
agent fill structured data in the plan. They get treated as fillable
structures, indexed and matchable.

### SVGs stay as assets with structural bonus metadata

SVG is partially auto-introspectable: aspect, palette (parsed from
fills/strokes), recolorability (explicit colors vs theme refs).
Semantic content needs vision — same describe path as raster assets.
The structural metadata is captured at ingest as bonus filter signal.

### One master per deck

PPTX allows multiple masters but it's vanishingly rare. v5 captures
one per deck. If a real deck breaks this assumption, add multi-master
support then. Customs / decoration-mixer / cross-deck decoration
grafting deferred to v5.x — the data model leaves room for it.

### Build from primitives, not deep-copy

`_copy_slide_into`, `_place_atom`, `_compose_custom_slide`, scheme +
font remap machinery in v4 — all gone in v5's build engine. The new
engine takes a chosen host theme + a list of plan entries; for each
plan entry it constructs a new slide on the host master and adds
primitive shapes (text boxes, tables, charts, pictures) at the
geometry the skeleton specifies. No XML grafting from foreign decks.

---

## Data model

### Skeleton (`templates/<id>/skeleton.yaml`)

Replaces v4's `meta.yaml` for slide entries. One file per skeleton.

```yaml
id: deckA_03
source_deck: deckA
source_slide_index: 3
categories: [data, comparison]    # multi-select from controlled enum
preview: preview.png              # thumbnail of source slide
slots:
  - id: title
    kind: heading
    geometry: {x: 0.05, y: 0.05, w: 0.9, h: 0.12}
    style:
      font_role: major            # +mj-lt — resolves to host theme's major font
      color_role: primary         # resolves to host theme's primary color
      size_pt: 36
      bold: true
      alignment: left
    constraints:
      max_chars: 60
      max_lines: 1
      required: true
  - id: left_body
    kind: bullets
    geometry: {x: 0.05, y: 0.25, w: 0.42, h: 0.65}
    style: {font_role: minor, color_role: text_default, size_pt: 18, alignment: left}
    constraints:
      max_items: 5
      max_chars_per_item: 80
      required: false
  - id: right_table
    kind: table
    geometry: {x: 0.55, y: 0.25, w: 0.42, h: 0.65}
    constraints:
      max_rows: 8
      max_cols: 4
      has_header: true
      required: false
```

**`kind` enum:** `heading`, `paragraph`, `bullets`, `table`, `chart`,
`image`, `footer`. Order in `slots:` array = reading order.

**`style.font_role` and `style.color_role`** are theme-relative tokens,
not concrete fonts/colors. They resolve against the chosen host theme
at build time — so a skeleton built on Acme's theme uses Acme's major
font and primary color even if the skeleton was originally extracted
from a different deck. This is how *font enforcement via host theme*
works: the skeleton never carries a concrete typeface, so there's
nothing to drift.

### Category enum (controlled, enterprise-flavored)

Initial set:

- `opening` — title, agenda, intro, kickoff
- `section_divider` — between major sections
- `content` — general body text + bullets, single topic
- `comparison` — 2-col side-by-side, vs, before/after
- `data` — table-heavy or chart-heavy
- `metric` — single large stat or KPI panel
- `quote` — pull-quote, testimonial, callout
- `closing` — Q&A, thank you, next steps, contact

Multi-select per skeleton. Auto-classifier (Python heuristic at
ingest) proposes a likely set based on slot inventory; user
accepts / edits in the Flask UI.

### Theme (`themes/<deck_id>/theme.yaml`)

Extends v4's `theme.yaml`. One per source deck.

```yaml
id: deckA
palette:
  primary: "#0A2540"
  accent: "#FF6B35"
  text_default: "#1A1A1A"
  background: "#FFFFFF"
  # …full theme color slots
fonts:
  major: "Aptos Display"
  minor: "Aptos"
master_pptx: master.pptx          # extracted host master fragment
preview: preview.png              # thumbnail of blank master layout
decorations:                      # informational, auto-classified
  - kind: top_bar
    geometry: {x: 0, y: 0, w: 1, h: 0.015}
    color_role: primary
  - kind: corner_logo
    geometry: {x: 0.92, y: 0.92, w: 0.06, h: 0.06}
    asset_id: asset_a1b2c3
  - kind: page_number
    geometry: {x: 0.92, y: 0.95, w: 0.06, h: 0.03}
```

In v5 decorations are not picked individually — they ride along with
the master verbatim. Inventory is captured and exposed informationally
so future v5.x can build a decoration-mixer (mix top-bar from deck A
with corner-logo from deck B on host C) without reingesting.

### Assets — unchanged describe path

Same vision-describe flow as today. SVGs get structural bonus metadata
(palette_hex, aspect, recolorable flag) added at ingest as filter
signal.

### Overlap and frozen backgrounds (experimental)

**Marked experimental.** This is a proposed handling for a real
digest-time problem. Auto-detection heuristic is decent but not
foolproof; if real-world usage shows this approach causes more
breakage than it prevents, remove the freeze-as-background logic and
fall back to auto-rejecting overlap-detected skeletons (user reviews
manually via the reject flow). When implementing in phase B, keep the
freeze-as-background module self-contained so removal is a small
revert, not a refactor.

**Problem.** A slide has a structural illustration (e.g. a chain of
4 squared shapes representing a process) with text labels overlaid
on each shape. In the structural-skeleton model:

- Dropping the picture (it's a freeform / auto-shape we'd conscious-
  drop) leaves the 4 text labels orphaned in nowhere positions —
  useless skeleton.
- Keeping the picture as an image slot lets the agent swap it for a
  random photo, which misaligns the labels — broken deck.

**Approach.** Freeze-as-background detection at digest time.

1. During digest, detect overlap: text shapes whose geometry
   overlaps a picture or freeform cluster's bounding box → the
   picture is *probably* a structural background illustration, not
   a swappable hero.
2. For those slides: render the non-slottable underlay (picture +
   freeforms + decorations) to a flat PNG, store as `background.png`
   on the skeleton, add `background_image: background.png` to
   `skeleton.yaml`. **Not a slot — a baked-in layer.**
3. Overlaid text shapes become normal slots with their original
   geometry.
4. Build engine paints the background first, then places filled
   slots on top.
5. Agent fills text slots; cannot swap the background. Matches
   structural intent — a process-flow skeleton works *because of*
   the chain image.
6. Flask UI surfaces detected overlaps; user confirms "freeze as
   background", overrides to "treat as image slot anyway", or
   rejects the skeleton.

**Schema impact.** Adds optional `background_image: <path>|null` to
`skeleton.yaml`. Build engine reads it; absent means no background
layer.

**Fallback if removed.** Strip steps 1-6. Replace with: at digest,
flag overlap-detected skeletons with `status: pending` + a
`digest_warnings: [overlap_detected]` field. User triages in Flask
UI — categorize anyway (accepting the broken-deck risk) or reject.

---

## Agent-facing methods (`reader.py`)

| Command | Purpose |
|---|---|
| `list-themes` | List available host themes with previews. |
| `list-skeletons --category <name> --has-slot <kind>` | Filter skeletons by category and slot inventory. |
| `get-skeleton <id>` | Full skeleton YAML + slot details. |
| `match-skeletons --content '<json>' [--category <name>]` | **Content-first match.** See below. |
| `list-assets --kind <kind> --tags …` | Asset pool query. |
| `get-asset <id>` | Asset description + binary path. |
| `validate-plan <plan.json>` | Full pre-build constraint check across all slides + asset choices. |
| `compose --theme <theme_id> <plan.json> <out.pptx>` | Build the deck on the chosen host theme. |

### Engine-side helpers — we compute, agent decides

Anything the agent could get wrong arithmetically, we expose as a
method. The agent's job is creative selection, not aspect-ratio math
or char counting in Unicode.

| Helper | Returns |
|---|---|
| `check-asset-fit <asset_id> <skeleton>.<slot>` | `{fits: true, will_resize_to: [W,H], will_crop: <region>\|none}` or `{fits: false, reason, suggestion}`. Covers aspect mismatch, min-resolution, kind mismatch (e.g. icon picked for hero slot). |
| `measure-text <str\|array> [--against <skeleton>.<slot>]` | Char count, word count, estimated line count at the slot's font size. With `--against`: pass/fail vs the constraint with current headroom. |

Plus **build-engine policy** the agent never calls (engine handles
silently):

- **Image auto-fit.** Default `cover` — center-crop preserving
  aspect to fill the slot. Slot can declare `auto_fit: cover |
  contain | stretch` to override. The agent never sees EMU
  coordinates or computes aspect ratios.
- **Text auto-wrap.** Within slot geometry. Vertical overflow is
  the trigger for `auto_shrink` (see overflow policy below) when
  the agent has signalled fall-through.

### `match-skeletons` — the content-first selection API

The agent describes the content it wants to put on a slide; the engine
returns the skeletons whose slots actually fit, ranked.

```bash
python reader.py match-skeletons --content '{
  "title": "Q4 results beat consensus",
  "bullets": [
    "Revenue +12% YoY to $1.8M",
    "Operating margin expanded 200bps",
    "Free cash flow positive for full year"
  ]
}' --category data
```

Returns:

```json
{
  "matches": [
    {
      "skeleton_id": "deckA_03",
      "categories": ["data"],
      "fit_score": 1.0,
      "slot_mapping": {"title": "title", "bullets": "body_bullets"},
      "headroom": {"title": "20 chars to spare", "bullets": "2 items to spare"}
    },
    {
      "skeleton_id": "deckB_07",
      "categories": ["data", "opening"],
      "fit_score": 0.85,
      "slot_mapping": {"title": "header", "bullets": "left_body"},
      "headroom": {"title": "5 chars to spare", "bullets": "1 item to spare"}
    }
  ],
  "issues": []
}
```

**`fit_score`** factors:
- Gate: does it fit at all (binary).
- Tightness: closer to slot constraint scores higher — the slide was
  *designed* for that length; loose-fit wastes the layout.
- Category match.
- Presence of optional slots the agent requested (e.g. agent passes
  `--has-slot footer` and skeleton has one).

**Zero-match drives the rephrase loop:**

```json
{
  "matches": [],
  "issues": [
    {
      "slot": "title",
      "your_value": "Q4 results beat consensus expectations across all geographic segments",
      "your_length": 73,
      "tightest_constraint": 60,
      "suggested_action": "rephrase to ≤60 chars (drop 13)"
    },
    {
      "slot": "bullets",
      "your_count": 7,
      "tightest_constraint": 5,
      "suggested_action": "consolidate to ≤5 items"
    }
  ]
}
```

### Rephrase loop with fall-through (the SKILL.md contract)

On `matches: []`, the agent's contract is a three-step fallback:

1. **Rephrase first.** Apply `suggested_action` (shorten title to ≤N
   chars, consolidate to ≤N bullets). Re-call `match-skeletons`.
2. **If rephrasing would lose meaning** — text is already terse,
   trimming further destroys content — use the text as-is and pass
   `overflow: "shrink"` in the plan value:
   `{"value": "...", "overflow": "shrink"}`. Build engine will
   auto-shrink font to fit and emit a warning to
   `<output>.warnings.json` for the user to fix manually after the
   deck is built.
3. **Picking a near-miss skeleton is not an option.** The escape
   hatch is `overflow: "shrink"` on the *intended* skeleton, not
   selection of a different one whose constraints don't match the
   intent.

`overflow: "shrink"` is an escape hatch, not the default. Constraints
exist because the source slide was designed for that length;
overflow degrades the visual. The warnings sidecar exists so the
user knows what to fix, not so the agent can ignore constraints
freely.

**Warnings sidecar shape (`<out>.warnings.json`):**

```json
{
  "warnings": [
    {
      "slide_index": 3,
      "skeleton_id": "deckA_07",
      "slot_id": "title",
      "constraint": {"max_chars": 60},
      "actual": {"chars": 73},
      "action_taken": "auto-shrunk font from 36pt to 30pt",
      "agent_note": "could not shorten without losing meaning"
    }
  ]
}
```

### `validate-plan` — pre-build safety net

After the agent has assembled a full plan across multiple slides
(mappings + image picks + table data + chart data), one shot check
before binary generation. Same constraint engine, operates on a full
plan. Catches anything `match-skeletons` couldn't catch upstream
(e.g., picked asset doesn't satisfy a slot's aspect constraint).

### How fonts and colors resolve at build time

- A slot's `style.font_role: major` resolves to the chosen host
  theme's `fonts.major`. Same for `minor`. Skeletons never carry a
  concrete typeface; nothing to drift across decks.
- A slot's `style.color_role: primary` resolves to the host theme's
  `palette.primary`. Same for `accent`, `text_default`, etc.
- A skeleton can override at slot level by specifying a concrete
  value (`color: "#FF0000"`) — escape hatch for the rare deliberate
  off-theme highlight.
- The plan can override at fill time per slot (`{"value": "...",
  "color": "#FF0000"}`) — escape hatch for the agent.

### Page numbers

The host theme's master carries a page-number placeholder. The build
engine preserves it on each generated slide and python-pptx
auto-sequences across the output deck. Source-deck page numbers are
not copied (would be wrong on a freshly-composed deck).

---

## Authoring workflow

```
1. cli.py ingest <deck.pptx>
   → workspace/themes/<deck>/
       theme.yaml          master + palette + fonts + auto-classified decorations
       master.pptx         extracted host master fragment (one master per deck)
       preview.png         master thumbnail
   → workspace/skeletons/<deck>_<NN>/
       skeleton.yaml       auto-digested: geometry + kind + style + heuristic constraints
       preview.png         thumbnail of source slide
       categories: pending
   → workspace/assets/<sha>.{ext|yaml}      unchanged from v4

2. app.py    (Flask describe UI, extended)
   New panel: "Categorize skeletons" — thumbnail grid; user assigns
              one or more categories per skeleton (auto-proposed set
              prefilled).
   New panel: "Verify theme" — master preview + decoration inventory;
              user accepts / edits decoration list, names theme.
   Existing panels: describe assets — unchanged.
   Removed: describe-slide (no longer a vision step).

3. cli.py validate
   - Promote categorized + theme-verified skeletons + complete assets
     to `done`.
   - Constraint sanity: a slot with required=true but no detected text
     in the source slide is flagged for human review.

4. cli.py build
   → dist/skill.zip
       themes/<id>/{theme.yaml, master.pptx, preview.png}
       templates/<id>/{skeleton.yaml, preview.png, background.png?}
       assets/<id>.{ext, yaml}
       reader.py
       SKILL.md
       index.json
```

### Skeleton status lifecycle

`skeleton.yaml` carries a `status:` field:

- `pending` — written by ingest. Awaiting category assignment (and
  awaiting overlap-handling decision if `digest_warnings` is set).
- `done` — categorized; included in build, listed and matchable.
- `rejected` — explicitly rejected by user via the Flask UI
  "Reject" button. Stays in workspace (re-ingest preserves the
  state — no auto-recategorize). Excluded from `validate`
  promotion, `list-skeletons`, `match-skeletons`, and build output.

**Rejection is reversible.** The Flask UI shows rejected skeletons
in a separate "Rejected" filter; user can click "Restore" to flip
status back to `pending` (returns to the categorize queue) or
straight to `done` if previously categorized (one-click un-reject
without re-categorizing).

Use cases for rejection: title-placeholder slides, empty
boilerplate, slides with overlap too messy to background-freeze
cleanly, anything visually useless. Use cases for un-reject: changed
mind after seeing the rest of the deck, mis-clicked, re-evaluating
after a brief discussion.

---

## Build phases

End-to-end-first. Each phase ships a working slice. v4 stays
operational until phase F flips the build flag.

| Phase | Touches | Deliverable |
|---|---|---|
| A | `REDESIGN.md`, `authoring/schemas/skeleton.yaml`, extended `theme.yaml` | Schemas + this doc. **You are here.** |
| B | `authoring/cli.py:ingest` | Digest pass: slot inventory with geometry, style, constraints. Auto-classify category. Capture `master.pptx` per deck. Auto-classify decorations. Overlap detection + freeze-as-background rendering (experimental — self-contained module so removable). |
| C | `authoring/app.py` | Categorize-skeletons panel + verify-theme panel. Reject/Restore controls (rejection is reversible). Promote on save. Remove describe-slide path. |
| D | `consumer/reader.py` | `list-themes`, `list-skeletons`, `get-skeleton`, `match-skeletons`, `validate-plan`. Tested in isolation against a fixture skeleton set. |
| E | `consumer/reader.py:compose` | New build engine: build slides from primitives on chosen host theme. All kinds (heading, paragraph, bullets, table, chart, image). Page numbers auto-sequence. Font / color role resolution. |
| F | `consumer/SKILL.md`, `authoring/cli.py:build` | Rewritten agent contract. v4 paths removed from build output. |
| G | `tests/test_v5.py` | Constraint engine, match ranking, build round-trip, theme application, rephrase loop sanity. |

Per-phase commits with multi-paragraph bodies — project convention
from v3/v4.

---

## Compatibility with v4

None at the artifact level. v5's `skill.zip` is a different shape; v4
readers won't open it, v5 reader won't open v4 bundles. **No
backwards-compat shim** — v4 plans assume slide-as-template-XML, which
v5 doesn't have.

For users with existing v4 workspaces, v5 ingest is re-run against the
same source decks. Structural extraction is deterministic, so
re-ingesting produces a stable set of skeletons. Existing asset
descriptions are reusable (assets are content-addressable by SHA).

---

## Open questions deferred

- **Auto-classifier propose vs cold assign.** Probably propose
  (heuristic: "has table → suggest data"; "very large heading + 1 line
  subtext → suggest opening"). Ship propose; let user override.
- **Partial-match returns in `match-skeletons`.** When zero exact
  matches, should we return near-matches flagged with which slots
  failed, so the agent can choose between rephrasing or relaxing
  category filter? Defer; ship strict-only first, add if real usage
  shows the rephrase loop spinning.
- **Decoration-mixer.** Deferred to v5.x. Theme data model already
  supports it (decorations carry geometry + asset_id).
- **Custom user-supplied master.** Deferred. User flagged this as
  desired for the future; v5 ships one-master-per-source-deck only.
- **Multi-deck output composition.** Explicitly disallowed in v5: one
  host theme per output deck.

---

## What this does NOT replace

- The Flask `/compose` UI (plan-builder + plan-runner) stays. Just
  the underlying engine changes.
- The asset describe-via-prompt flow stays unchanged.
- The two-layer split (authoring repo vs portable consumer zip) stays.
- `theme.yaml` per deck (from v4) stays and is extended.

---

## Pointers for next session

- v4 codebase is `main` head. The `fix/multi-slide-compose-rels`
  branch (rels-fix for chart placement) is moot under v5 and is not
  merged into this redesign branch.
- v4 [`HANDOVER.md`](HANDOVER.md) describes the v4 architecture and
  the v4.1 punch list. Items 1 (chart/SmartArt placement) and 3
  (aspect scaling) are obviated by v5. Item 2 (font remap) and item 4
  (group recursion) are already-landed and don't carry over (font is
  now role-based; group recursion is not relevant when shapes are
  built from primitives, not deep-copied).
- v4 [`FINDINGS.md`](FINDINGS.md) audit items: most of part B
  (colors / SVG / non-photo / two-way YAML) is addressed by v5's
  structural model. Part A items survive piecemeal — re-triage in
  context of v5.
- All v5 work lands on `redesign/structural-slides`. No premature
  merges to `main`; v5 is a directional pivot and ships as a clean
  cut once phases A-G are end-to-end working.
