# pptx-skill — findings for next session

Two deliverables from a remote walk-through, written down so a local
session tomorrow can pick them up. Nothing in here is implemented yet.

- **Part A** — code/architecture audit: what could be done better.
- **Part B** — verification of how the skill would behave for
  non-photo assets (SVG, vector, callouts), colors, text formatting,
  and whether the YAML/plan can round-trip that info both ways.

All file/line references are against the current branch
(`claude/pptx-skill-repo-access-LrLDr`, base 5bfe2e9).

---

# Part A — Walk-through audit (what could be done better)

## A1. Architecture / design

1. **`brand.md` is authoring-only.** `compose_page` injects it into
   prompt bundles, but `build` does NOT ship it inside `skill.zip`.
   Any agent consuming the zip directly (the whole point of the
   consumer split) never sees brand rules. Options: (a) embed
   brand.md into the zip, or (b) document explicitly that brand only
   applies when prompt bundles are generated via the local UI.

2. **Three duplicate copies of the controlled vocab.** Enums for
   `feel` / `suitable_for` / `kind` / `composition` exist in
   `cli.py` (Python), `app.py`'s inline HTML (JS), and the prompt
   markdown files. Drift is inevitable. SKILL.md lists *merged*
   feels (`formal, punchy, data-dense, warm, clinical, minimal,
   playful, dramatic, celebratory`) even though slide-feels and
   asset-feels are different sets — an agent filtering templates by
   `feel=playful` will silently get zero results.

3. **First-template-as-host compose strategy**
   (`reader.py` `cmd_compose`, ~lines 410-447). All later slides are
   deep-copied onto a "blank-ish" layout from the host's master. If
   the user mixes templates from decks with different
   masters/themes/aspect ratios, you get the first deck's master
   applied to *foreign* shape trees → font/color/theme
   inconsistencies. No way to override the host or check
   compatibility.

4. **`_copy_slide_into` strips layout placeholders then deep-copies
   XML.** `reader.py:299-301` removes all shapes the chosen layout
   brought along — meaning master-level decoration (page numbers,
   brand bars) is lost on every appended slide. The "pick the
   layout with the fewest placeholders" heuristic at
   `reader.py:286-295` is arbitrary and not deterministic across
   decks.

5. **`prs.part.drop_rel(rId)` swallowed with bare `except`**
   (`cli.py:278-280`). Silently leaks slide relationships into
   fragments → bloated `.pptx` files (each fragment carries the
   whole original deck's layouts/masters/media even though only one
   slide is exposed).

6. **No tests.** Not a single test file. For something that mutates
   user `.pptx` files via XML surgery, regression cover would catch
   a lot.

## A2. Ingest layer (`cli.py`)

7. **`PP_PLACEHOLDER.OBJECT` treated as textual**
   (`cli.py:110-116`). OBJECT placeholders can hold tables, charts,
   or embedded objects, not just text. The current code creates a
   `text`/`bullets` slot, then compose tries to fill it with a
   string — undefined behavior on a chart placeholder.

8. **`position_quadrant` uses top-left, not center**, but variables
   are named `cx`/`cy` with a comment "we use top-left for
   orientation" (`cli.py:86-102`). Misleading naming; a wide
   centered shape near `top=0` gets classified `top-left` even
   though it's visually `top-center`.

9. **`aspect_ratio` returns `3:2`/`2:3` tokens**
   (`cli.py:75-77`) that aren't documented in schemas/prompts.

10. **20% area threshold for "this is a slot" is hardcoded**
    (`cli.py:221`). No way to tune per-deck. A 19%-area photo
    silently becomes frozen background.

11. **Slot id ordering is shape-iteration-dependent.**
    `detect_slots` iterates `slide.placeholders` then
    `slide.shapes`. Re-ingest of the same deck is currently stable,
    but any future reorder in python-pptx would silently rename
    slots → existing YAMLs become disconnected.

12. **Wasteful slide fragments.** Each `slide_NN.pptx` is the
    *whole deck* with N-1 slides removed but all the
    masters/layouts/media still embedded. For a 50-slide deck this
    is ~50× duplication on disk.

13. **No way to remove a slide from the workspace.** Re-ingest is
    idempotent but if you delete a slide from the source deck and
    re-ingest, the old `slide_NN.yaml` lingers.

14. **`asset_id = f"asset_{sha[:8]}"`** (`cli.py:361`). 8 hex chars
    = 32 bits → birthday collision around ~65k assets. Probably
    fine forever; worth noting.

## A3. Compose layer (`reader.py`)

15. **Font color is *not* copied** in `_fill_text_shape`
    (`reader.py:128-161`) or `_fill_bullets_shape`
    (`reader.py:163-200`). Size/bold/italic/name are copied; color
    is omitted. If your template has a brand-color title, after
    compose it falls back to default black/theme-default.

16. **`_replace_image_shape`** (`reader.py:202-219`) destroys the
    original shape and re-adds with `add_picture`. Loses: crop,
    picture borders, rotation, shadow, transparency, picture
    effects, alt text. For decks where photos have a brand-styled
    crop or border, this strips it.

17. **Bullets' multi-line-within-a-value handling**
    (`reader.py:197-199`) appends extra lines as `"\n" + extra`
    runs inside the same paragraph. PPTX treats `\n` in a run as a
    soft line break but font copy semantics differ between sub-runs
    and primary runs — visual results vary by template.

18. **`compose` shells out via subprocess** in `app.py:917-952`
    instead of importing `reader.py` directly. 60s timeout,
    requires the file to exist on disk, makes error reporting
    awkward.

19. **`matches_filter` requires the key to be present**
    (`reader.py:88-91`). Filtering by `feel=warm` excludes items
    with no `feel` at all, which may be fine — but it's not
    documented and there's no "include unset" mode.

## A4. Describe / batch flow (`app.py`)

20. **Batch prompts use YAML** (`app.py:246-304`). The
    `_recover_flat_batch_yaml` heuristic at `app.py:310-353` is a
    band-aid for LLMs that get indentation wrong. The
    fragile-format smell is loud — multiple `CRITICAL` /
    `INCORRECT` / `DO NOT` blocks in the bulk instructions. **JSON
    would eliminate the entire failure mode** (no significant
    whitespace). The per-item schema is small enough; YAML brings
    nothing here except parser fragility.

21. **Batch limit is hardcoded to 20** (`app.py:361`). No way to
    override.

22. **`_ensure_slide_png` requires macOS `qlmanage`**
    (`app.py:84-106`). Linux/Windows users can't use the
    bulk-slide batch flow. `cli.py preview` uses LibreOffice for
    the same job — these could share one renderer.

23. **`PRESET_NAME_RE` and `BATCHES_DIR` are workspace-internal
    but undocumented.** A user looking at `workspace/` sees
    mysterious `_presets/`, `_batches/`, `_compose_out/`
    directories without explanation.

24. **Inline HTML/CSS/JS in `app.py`** — 1,200+ lines of template
    string. Standard early-stage tradeoff, but worth noting for
    future maintainability.

## A5. Schema / data model

25. **No structured palette / theme on slides.** `colors` exists
    only on assets, and as color *names* (`navy`, `warm white`),
    not hex. Templates have no way to declare "this slide uses
    primary-on-secondary" — an agent matching brand colors has no
    signal beyond the asset.

26. **`scope` exists on assets but not on slides.** A slide that's
    heavy on a specific client's logo or program-specific
    terminology can't be tagged as `client:acme-bank`. Agents will
    pick it for any deck.

27. **`depicts` is asset-only.** A slide that's *specifically*
    about workflow-X has no equivalent field — only the free-text
    `intent`.

28. **`max_chars` / `max_items` are *hints* not constraints.**
    Validate-time and compose-time both ignore them. An LLM that
    produces a 200-char title for a 60-char slot silently
    overflows.

29. **`feel` enum split between slide (`data-dense`,
    `celebratory`) and asset (`minimal`, `playful`, `dramatic`) is
    opaque.** Likely intentional but worth documenting why; right
    now it just looks like an oversight.

## A6. Docs / DX

30. **`README.md` quickstart says
    `pip install -r authoring/requirements.txt`** but doesn't
    mention LibreOffice or qlmanage prerequisites for
    preview/batch.

31. **`PLAN.md` is the design doc, but it doesn't match
    implementation in places** (e.g., describes 7 commands but
    `package-app` makes it 8; doesn't mention the Flask UI /
    compose page / batch flow / brand.md / presets, all of which
    are major surfaces).

32. **No CHANGELOG / version pin on the zip artifact.** Consumer
    agents can't tell which authoring version produced the
    bundle.

---

# Part B — Colors / SVG / non-photo assets / two-way YAML

Question being answered: today, would the skill correctly handle a
scraped presentation that includes **colors, SVGs, and other
non-raster assets**, and could the agent's JSON plan and the YAML
round-trip carry color and text-formatting info? Short answer: **no
on all three counts, but the gaps are diagnosable and the
architecture is fixable without a rewrite.**

## B1. What ingest currently captures (the source of the limitation)

`extract_picture_assets` in `cli.py:345-368` iterates
`slide.shapes` and filters with `_shape_is_picture` (`shape_type ==
MSO_SHAPE_TYPE.PICTURE`). Everything else is ignored as a
discoverable asset.

| Shape kind                                | Captured as asset? | Survives as slide XML? | Slot-addressable? |
|---|---|---|---|
| Raster `Picture` (PNG/JPG)                | yes                | yes                    | yes |
| SVG `Picture` (Office stores w/ raster fallback) | partial — see B2 | yes              | yes (as the fallback raster) |
| EMF / WMF vector picture                  | yes (as `.emf`/`.wmf` blob) | yes           | partial — re-render is fragile |
| `AUTO_SHAPE` (rectangles, arrows, callouts) | no               | yes                    | no |
| `FREEFORM` (custom paths)                 | no                 | yes                    | no |
| `GROUP`                                   | no                 | yes                    | no |
| `LINE`, connectors                        | no                 | yes                    | no |
| `CHART`                                   | no                 | yes                    | no (treated as OBJECT placeholder, currently mis-typed as text) |
| `TABLE`                                   | no                 | yes                    | no |
| SmartArt                                  | no                 | partial                | no |
| Theme/master backgrounds                  | no                 | yes (via master)       | no |

Vector/structural assets ride along inside the slide fragment XML
(deep-copied at compose time), but they're invisible to the index,
unfilterable, and can't be swapped or recolored from a plan.

## B2. SVG specifically

PPTX stores SVG as a `Picture` shape with **both** the raw SVG and
a raster fallback (PNG or EMF) embedded as separate image parts.
python-pptx's `shape.image` returns the *primary* part — which is
typically the raster fallback, not the SVG. So today's
`extract_picture_assets`:

- Picks the PNG/EMF fallback.
- Writes it to disk as `<sha1>.png` (or `.emf`).
- Loses the SVG entirely.

Compose-time `_replace_image_shape` then `add_picture`s the raster
fallback. The visual result is acceptable, but you've thrown away
the vector source — bad for re-coloring, scaling, or
DPI-independent rendering.

To actually keep the SVG you'd need to dig into the shape's
relationship XML and pull both parts. python-pptx doesn't expose
this cleanly; you'd be working with `shape._element.xpath(...)`
against the `asvg:svgBlip` attribute.

## B3. Colors — can YAML / plan / compose carry them?

**Today: no, in three places.**

1. **Schema doesn't model colors structurally.**
   - `asset.yaml` has `colors: []` but it's *names* (`navy`, `warm
     white`) used for retrieval, not hex.
   - `slide.yaml` has no `colors` field at all.
   - No `palette`, `theme`, `tokens` block anywhere. `brand.md` is
     free-text prose.

2. **Plan JSON has no color slot.**
   - `cmd_compose` only understands three value shapes: bare string
     (text), array of strings (bullets), `asset_<id>` string
     (image). See `_apply_slot_value` at `reader.py:222-264`.
     There's no `{"color": "#0A2540"}` branch.

3. **Compose drops color even when the template had it.**
   - `_fill_text_shape` copies `font.size / bold / italic / name` —
     explicitly *not* `font.color`. See `reader.py:148-156`. So if
     you ingest a slide with a navy title and compose with a new
     title string, you get default-color text.

**What would it take?**

Doable without a rewrite. Three coordinated changes:

- **Schema**: add `palette: {primary: "#...", accent: "#...", ...}`
  to brand.md (structured YAML, not prose); add an optional
  `colors: {hex: "...", role: "primary|accent|..."}` on assets;
  add `theme_colors` to slide.yaml extracted at ingest from
  `master.color_theme`.
- **Plan**: extend slot value to be either a string OR
  `{"text": "...", "color": "#...", "bold": true, "runs": [...]}`.
  Reader needs a small dispatcher.
- **Compose**: in `_fill_text_shape`, default to copying
  `font.color` from the template run; let the plan override.

Mostly additive — existing string-valued plans keep working.

## B4. Text formatting (bold / italic / runs / alignment)

Plan can only express **a single style per slot**, derived from
the template's first run. There is no way for the agent to say:

- "Make the word `risk` red and bold in this title."
- "Right-align this subtitle."
- "Use the brand accent color for the lead number."

The fix is additive — accept a `runs: [...]` array in slot values
where each run carries its own style — but it'd shift the plan
format from "dumb strings" to "rich text deltas", which is a
meaningful complexity bump for the consuming agent.

## B5. Two-way YAML flow

Today the YAML flow is **one-way**: ingest writes structural
fields → human/LLM fills descriptive fields → validate freezes
them → build packages them → compose reads them. There is no path
where the agent (or anything downstream) writes back to YAML.

This matters because if you want colors/formatting to be
**collected from the source deck**, that's a *new ingest pass*
enriching the YAML with extracted theme/palette/run-formatting
data. python-pptx exposes enough to do it:

- `slide.slide_layout.slide_master.element.xpath('.//a:clrScheme')`
  → theme palette (6 colors + 2 accents).
- For each run in the template: `run.font.color.rgb`,
  `run.font.color.theme_color`, `run.font.name`,
  `run.font.size`.
- Asset binaries are already SHA1'd; you could add a sidecar field
  `dominant_colors_hex: [...]` via a quick raster sample (PIL
  `Image.quantize`) without needing the LLM.

If you ALSO want the agent's plan output to feed back into YAML
(e.g., to learn "this template gets used for risk slides 80% of
the time"), that's a separate observability loop — not currently
wired in.

## B6. Concrete verdict by scenario

| Scenario                                                    | Today? | Why |
|---|---|---|
| Scrape a deck with brand-color titles, regenerate keeping the brand color | no | font.color not copied at compose |
| Tell the agent "use the brand accent for this title"        | no | Plan has no color field; no structured palette |
| Capture an SVG logo from the deck and re-emit at any size   | partial | SVG falls back to raster on ingest |
| Capture a SmartArt diagram so the agent can pick it         | no | Not extracted as asset |
| Capture a callout shape (e.g. an arrow) for reuse           | no | Not extracted |
| Carry chart data as a swappable slot                        | no | OBJECT placeholder mis-typed as text |
| Filter assets by hex / role rather than color name          | no | Schema only has names |
| Round-trip palette: deck → YAML → deck                      | no | No palette block; brand.md is prose |

---

# Suggested order of attack (for tomorrow's local session)

If the goal is "make non-photo and color-aware decks work
end-to-end" the lowest-friction path is:

1. **Fix the color regression first** (A15) — copy `font.color`
   in `_fill_text_shape` / `_fill_bullets_shape`. One-line
   defensive change, immediately visible improvement, no schema
   churn.

2. **Add structured palette extraction at ingest** (B5 bullet 1)
   — read `clrScheme` from the master, write it into a new
   `theme.yaml` per deck or into each `slide.yaml`. Read-only on
   the schema side; nothing else has to change yet.

3. **Decide the plan-format extension** (B3 / B4) — the
   string-or-object slot-value shape. This is the architectural
   call that everything else hangs off; worth a whiteboard before
   typing.

4. **Then** revisit SVG extraction and non-picture asset capture
   (B1, B2). These are bigger lifts and benefit from the schema
   work being settled first.

5. Cleanup items in A1-A6 can be triaged independently as you
   go — none of them block the above.
