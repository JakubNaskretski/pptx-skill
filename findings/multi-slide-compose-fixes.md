# Handover — Multi-slide compose: rels corruption + background loss

**Status**: pre-implementation. Two bugs in `consumer/reader.py` cause
broken output when composing 2+ slides. Both root-caused and
reproducible. This document is the spec.

**Decided scope**: full fix for both bugs (the ambitious option in each
case). They share architecture: once slides are *flattened* into
self-contained units, the graft step is "copy shapes with proper
transitive rel import" — the same machinery resolves both problems.

**HEAD at planning time**: `e04dbde`.

---

## TL;DR for the picking-up engineer

Compose a 1-slide plan → output is mostly fine but loses backgrounds
defined on the source's layout/master. Compose a 2+ slide plan →
output triggers PowerPoint's "couldn't read some content, removed it"
repair dialog and is missing shapes. Both come from a too-shallow
slide-copy implementation in [`consumer/reader.py:_copy_slide_into()`
(lines ~695-772)](../consumer/reader.py).

Fix is in two architectural pieces, deliverable in three shippable
phases. Each phase has a clear acceptance test.

---

## Repro — anchor your mental model

Reproduces on any non-trivial ingested deck. Run these against
whatever's already in `authoring/workspace/decks/`:

```bash
# Start the app if not running
python3 authoring/app.py &
sleep 2

# Pick the first ingested deck's first 3 template ids
IDS=$(python3 -c "
import yaml; from pathlib import Path
for deck in sorted(Path('authoring/workspace/decks').iterdir()):
    ys = sorted((deck / 'slides').glob('slide_*.yaml'))[:3]
    if len(ys) >= 3:
        for y in ys:
            print(yaml.safe_load(y.read_text())['id'])
        break
")
read -r T1 T2 T3 <<< $(echo "$IDS")

# Compose 1 slide → no warning on open, but background may be missing
curl -s -X POST http://127.0.0.1:5050/api/compose/run \
  -H "Content-Type: application/json" \
  -d "{\"plan\": [{\"template\":\"$T1\",\"slots\":{}}]}" \
  -o /tmp/single.pptx

# Compose 3 slides → triggers PowerPoint's repair dialog on open
curl -s -X POST http://127.0.0.1:5050/api/compose/run \
  -H "Content-Type: application/json" \
  -d "{\"plan\": [
    {\"template\":\"$T1\",\"slots\":{}},
    {\"template\":\"$T2\",\"slots\":{}},
    {\"template\":\"$T3\",\"slots\":{}}
  ]}" \
  -o /tmp/multi.pptx

# Verify the corruption signature — Bug 1's smoking gun
rm -rf /tmp/inspect && mkdir /tmp/inspect && cd /tmp/inspect
unzip -q /tmp/multi.pptx
for sx in ppt/slides/slide*.xml; do
  sn=$(basename "$sx")
  used=$(grep -oE '(r:embed|r:link|r:id)="rId[0-9]+"' "$sx" | grep -oE 'rId[0-9]+' | sort -u)
  have=$(grep -oE 'Id="rId[0-9]+"' "ppt/slides/_rels/${sn}.rels" 2>/dev/null \
         | grep -oE 'rId[0-9]+' | sort -u)
  echo "$sn missing: $(comm -23 <(echo "$used") <(echo "$have") | tr '\n' ' ')"
done
```

Expected output on broken code:
- `slide1.xml missing:` (empty — host slide is untouched)
- `slide2.xml missing: rId<N>`
- `slide3.xml missing: rId<N>`

On the fixed code: every line ends with `missing:` (nothing).

---

## Codebase orientation (5 minutes)

- `authoring/` is build-time. Not in the compose path. Ignore unless
  you need to re-ingest.
- `consumer/reader.py` is the engine. **The bugs live here.**
  - Entry: `cmd_compose()` around line 1180.
  - Hot spot: `_copy_slide_into()` at ~695. Two issues, same function.
- `authoring/workspace/decks/<deck>/slides/slide_NN.pptx` — single-slide
  fragments produced by `cli.py:write_slide_fragment()` at ingest.
  Each one is a self-contained .pptx; the compose flow loads these
  fragments and grafts them together.
- `consumer/helpers/` — agent-side utilities, irrelevant to this fix.

`python-pptx` is the OOXML library. Its `Presentation.part` exposes
the underlying package; rels manipulation lives on each part via
`.rels`, `.relate_to(target_part, reltype)`, and similar. Read
[python-pptx's "Working with relationships" section](https://python-pptx.readthedocs.io/)
once before starting — it's short and clarifies the part/rel model
this fix depends on.

A `.pptx` is a zip with this structure:

```
ppt/slides/slide1.xml                       ← shape tree, references resources by rId
ppt/slides/_rels/slide1.xml.rels            ← maps rIds to actual parts
ppt/media/image1.png                        ← actual binaries
ppt/slideLayouts/slideLayout1.xml           ← layout templates
ppt/slideMasters/slideMaster1.xml           ← master (background inheritance lives here)
ppt/theme/theme1.xml                        ← colors, fonts, bgFillStyleLst
```

Each `slide<N>.xml.rels` is a *file-local* namespace — `rId2` in slide1's
rels and `rId2` in slide2's rels may point at completely different
things. This locality is what trips up the deepcopy in `_copy_slide_into`.

---

## Bug 1 — Multi-slide rels corruption (data loss)

### Symptom
Slides 2..N in the compose output contain shape XML referencing rIds
that don't exist in their `.rels` files. PowerPoint repairs by removing
the shape and shows the "couldn't read some content" dialog.

### Root cause
`_copy_slide_into()` at [`consumer/reader.py:740-770`](../consumer/reader.py#L740-L770)
special-cases pictures (calls `add_picture(BytesIO(blob), ...)` which
python-pptx implements correctly), then falls through to:

```python
el = copy.deepcopy(shape._element)
# ... optional color/font remap ...
new_slide.shapes._spTree.append(el)
```

This copies the shape XML — including any rId-bearing attributes
(`r:embed`, `r:link`, `r:id`) — and appends to the destination's
shape tree. It never imports the underlying part the rId pointed at,
never registers a new rel in the destination's rels file, and never
rewrites the rId in the copied XML.

Shape kinds that trigger this in practice:
- Auto-shapes with picture-fill
- Group shapes containing pictures or links
- Shapes with hyperlinks (`a:hlinkClick`, `a:hlinkHover`)
- Text-runs with hyperlinks
- Embedded charts and SmartArt (`p:graphicFrame` with chart or
  diagram payload)
- Tables with embedded media in cells (rare but exists)
- OLE embed shapes (`p:oleObj`)

### Fix specification

Replace the deepcopy branch with a rel-aware import. Pseudo-code:

```python
# In consumer/reader.py, alongside _copy_slide_into:

# Attributes that carry rId references in OOXML.
_REL_ATTRS = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed",
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link",
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
)

def _iter_rel_attr_holders(el):
    """Yield (element, attr_qname) for every rId-bearing attribute in the
    subtree rooted at el."""
    for descendant in el.iter():
        for q in _REL_ATTRS:
            if q in descendant.attrib:
                yield descendant, q

def _import_rel(src_part, dest_part, old_rid):
    """Look up old_rid in src_part's rels, import the target into
    dest_part's package via relate_to, return the new rId.

    For TRANSITIVE imports (chart/SmartArt payloads): recursively walk
    the target_part's own rels and import its dependents first, then
    register the top-level part. python-pptx's relate_to handles the
    common case; we wrap it to handle the transitive walk."""
    src_rel = src_part.rels.get(old_rid)
    if src_rel is None:
        return None
    target = src_rel.target_part
    # If target has its own rels (chart → embedded xlsx → image), they
    # come along automatically when target_part is registered because
    # rels travel with the part. python-pptx packages parts plus their
    # rels graph atomically.
    new_rid = dest_part.relate_to(target, src_rel.reltype)
    return new_rid

def _import_shape_xml(src_slide_part, dest_slide_part, dest_sptree, src_shape_el):
    """Import a single shape, rewriting rIds to dest-local rels."""
    new_el = copy.deepcopy(src_shape_el)
    rid_subs = {}
    for el, attr_q in _iter_rel_attr_holders(new_el):
        old_rid = el.get(attr_q)
        if old_rid is None or old_rid in rid_subs:
            continue
        new_rid = _import_rel(src_slide_part, dest_slide_part, old_rid)
        if new_rid is not None:
            rid_subs[old_rid] = new_rid
    # Second pass: substitute the rIds.
    for el, attr_q in _iter_rel_attr_holders(new_el):
        old_rid = el.get(attr_q)
        if old_rid in rid_subs:
            el.set(attr_q, rid_subs[old_rid])
    dest_sptree.append(new_el)
```

Then in `_copy_slide_into`, replace the deepcopy-and-append branch:

```python
# REMOVE:
# el = copy.deepcopy(shape._element)
# ...
# new_slide.shapes._spTree.append(el)

# REPLACE with:
_import_shape_xml(
    src_slide_part=src_slide.part,
    dest_slide_part=new_slide.part,
    dest_sptree=new_slide.shapes._spTree,
    src_shape_el=shape._element,
)
# Then apply scheme/font remaps on the inserted element (do it after
# the rId substitution so the remap operates on the same tree).
```

### Transitive rel handling for charts/SmartArt

A chart's `p:graphicFrame` references a chart part via `r:id`; the chart
part references an embedded `.xlsx` via its own rels; the xlsx
references images via its rels. When `dest_part.relate_to(target_part,
reltype)` registers the target with the destination package,
**python-pptx brings the target's own rels along** (the rels file is
part of the part). So shallow `_import_rel` already handles transitive
deps in practice — but verify in testing.

SmartArt is similar: `p:graphicFrame` → diagram colors/styles/data parts
all linked via the diagram part's rels.

### Edge cases to handle

1. **Slides whose shapes reference parts that don't exist at source**
   (already broken at ingest). Don't crash — log a warning, skip the
   rId substitution for that attribute, let the broken ref propagate.

2. **Hyperlinks** (`a:hlinkClick`) where the target is an external URL,
   not an internal part. These also use rIds but the rel's reltype is
   `hyperlink` and the target is `External=true`. `relate_to` doesn't
   handle this cleanly; need a special path that creates an
   external-target rel.

3. **rId collisions** in nested groups. After import, the destination
   slide's rels file may have rIds that collide with rIds being
   imported from a *third* slide later in the plan. python-pptx
   generates fresh rIds on each `relate_to`, so this is handled — but
   verify.

### Acceptance

1. The repro script's `missing:` column is empty for every slide in
   the 3-slide compose output.
2. PowerPoint opens the multi-slide output without showing the repair
   dialog.
3. The regression test (below) passes.

### Effort + risk

- ~150-200 lines including helpers + tests.
- Medium risk on hyperlinks and OLE embeds (uncommon but possible to
  hit). Mitigation: log a warning when a rel-attr type is encountered
  that the import code doesn't recognize, instead of failing silently.
- Low risk on the common case (picture-fills, group shapes with
  pictures) — that's the dominant code path.

---

## Bug 2 — Background / layout loss

### Symptom
Grafted slides (everything after slide 1) inherit the destination's
"blank" layout's background, not their original layout's background.
Single-slide compose uses the host's first slide on its native layout,
but the fragment extraction may have already stripped that layout's
background. Net result: backgrounds appear or disappear inconsistently
depending on which deck is the host and how many slides are grafted.

### Root cause
Two contributing places:

1. `_copy_slide_into()` at lines 715-727 picks the destination's layout
   with the fewest placeholders as the parent for the new slide:

   ```python
   for layout in dest_prs.slide_layouts:
       count = len(layout.placeholders)
       if best_count is None or count < best_count:
           dest_layout = layout
   ```

   That layout's background applies to the grafted slide. The source's
   original layout is never copied across.

2. `cli.py:write_slide_fragment()` followed by
   `prune_unreachable_parts()` may strip parts that the inherited
   background style needs. Specifically, when the source deck uses
   `<p:bgRef idx="N">` referencing theme's `<a:bgFillStyleLst>` entry
   N, and that entry uses a `<a:blipFill>` referencing a media part,
   the closure walk in the pruner must reach that media. **Verify
   this works** before assuming the fragment is the problem.

### Fix specification — Option C: flatten before graft

We chose the most ambitious option: before grafting, *flatten* each
source slide so its background is self-contained. After flattening,
the slide has an explicit `<p:bg>` element with all needed
sub-elements and media inline, and no longer relies on layout/master
inheritance for backgrounds. The graft step from Bug 1 then handles
everything correctly.

#### Flatten algorithm

```python
def _flatten_slide_background(slide):
    """Make `slide`'s background self-contained by copying its inherited
    background into the slide itself, then resolving any references.

    Mutates slide.xml in place. Idempotent on already-flattened slides.
    """
    sld_el = slide._element
    cSld_el = _find_or_create(sld_el, "p:cSld")
    if _has_explicit_bg(cSld_el):
        return  # already flattened or always had its own bg
    
    # Walk inheritance chain to find the effective background:
    layout = slide.slide_layout
    master = layout.slide_master
    theme = _theme_for_master(master)
    
    bg_source = _find_effective_background(slide, layout, master, theme)
    if bg_source is None:
        return  # nothing to flatten (slide has no background defined anywhere)
    
    # Materialize the background:
    if bg_source.kind == "bgPr":
        # Direct background properties (solid/grad/pattern/blip fill).
        new_bg = copy.deepcopy(bg_source.element)
        # If it has a r:embed referring to a media part, import the
        # part into the slide's own rels (same machinery as Bug 1).
        _import_rels_in_subtree(
            src_part=bg_source.containing_part,
            dest_part=slide.part,
            subtree=new_bg,
        )
    elif bg_source.kind == "bgRef":
        # Theme-defined background style. Resolve idx → theme's
        # bgFillStyleLst[idx-1001] → expand to inline bgPr.
        resolved_fill = _resolve_theme_bg_ref(theme, bg_source.idx)
        new_bg = _build_explicit_bg_from_fill(resolved_fill)
        _import_rels_in_subtree(theme.part, slide.part, new_bg)
    
    _insert_bg_as_first_child(cSld_el, new_bg)
```

Helper details:

- `_find_effective_background` walks `slide → layout → master → theme`
  and returns the first non-None background definition along with
  enough metadata to materialize it.
- `_resolve_theme_bg_ref` reads the theme's `<a:fmtScheme>
  <a:bgFillStyleLst>` and picks the entry by `bgRef@idx`. The idx is
  1001-based per OOXML spec.
- `_import_rels_in_subtree` is the same logic as Bug 1's
  `_iter_rel_attr_holders` + `_import_rel`, but operates on whatever
  subtree you pass in (could be a shape, could be a bg block).

#### Where to call flatten

In `_copy_slide_into()`, *after* opening the source slide and *before*
the shape-copy loop:

```python
src_prs = Presentation(str(src_slide_pptx))
src_slide = src_prs.slides[0]
_flatten_slide_background(src_slide)        # ← new
# (rest of the function — pick dest layout, copy shapes, etc.)
```

The flattened background now lives on the source slide's `<p:cSld>`
element. When we use the cleaner Bug 1 fix to copy shapes, we ALSO
need to copy this `<p:bg>` element to the destination slide's
`<p:cSld>`. One more line.

### Acceptance

1. The grafted slides' backgrounds match the source slides'
   backgrounds (visual inspection — open in PowerPoint and compare to
   the original deck's matching slides).
2. The output `.pptx`'s `slide<N>.xml` files contain explicit `<p:bg>`
   elements (no longer relying on layout inheritance).
3. Single-slide compose preserves the slide's original background
   (because flatten also applies to the first slide if going through
   `_copy_slide_into` — or, equivalently, apply flatten at ingest in
   `write_slide_fragment` so all fragments are pre-flattened on disk).

### Effort + risk

- ~200-250 lines including the theme-bg-ref resolver. Bg resolution
  has many cases (solid, gradient with stops, blip with crop, pattern,
  grouped fills).
- Medium-high risk. OOXML's background model is one of the gnarlier
  bits of the spec. Edge cases: backgrounds with `<a:blipFill>` and a
  `srcRect`/`stretch` modifier; gradient stops referencing scheme
  colors (need to resolve scheme colors at flatten time too); tile
  patterns.
- Mitigation: implement solid + linear-gradient + simple blip-fill
  first; warn-and-skip more exotic forms. Add coverage incrementally.

---

## Phased delivery — three shippable milestones

### Phase 1: Bug 1 core (1 day)

Rel-aware shape import for the **common** non-picture kinds:
auto-shapes, text boxes, group shapes, shapes with `r:embed` /
`r:link` attrs. Warn (don't crash) on shape kinds the import code
hasn't covered yet. Skip charts and SmartArt with a clearer warning
than today's silent skip.

**Ship criteria**:
- Repro script's `missing:` column is empty.
- PowerPoint opens 3-slide multi-compose output without repair dialog.
- Regression test passes.

### Phase 2: Bug 1 transitive (~half day)

Verify charts and SmartArt come through with their nested rels
correctly. If python-pptx's `relate_to` handles the transitive walk
(it should — rels are part-bound), Phase 2 is just lifting the
"skip-with-warning" guard and adding tests. If it doesn't, add a
recursive rel-import.

**Ship criteria**:
- Compose plan that picks templates with chart and SmartArt atoms
  produces output that opens cleanly *with the chart/SmartArt
  rendered* (not just blank-removed).

### Phase 3: Bug 2 — flatten before graft (1.5 days)

Implement `_flatten_slide_background` and integrate it. Start with
solid + gradient + simple-blip fills. Add other fill types in
follow-ups as needed.

**Ship criteria**:
- Multi-slide output's grafted slides have the same backgrounds as
  their source slides in the original deck (visual inspection).
- Slide XML contains explicit `<p:bg>` blocks for each affected slide.

After Phase 3, also re-ingest one deck to verify fragments preserve
backgrounds correctly (Bug 2 has both an ingest-side and compose-side
component; we're addressing it compose-side, but fragment quality is
worth a quick check).

---

## Test strategy

Add `tests/test_multi_slide_compose.py` (new file):

```python
"""Regression tests for multi-slide compose.

Asserts the output .pptx has no broken rId references, opens cleanly
in python-pptx, and (for Phase 3) carries explicit background blocks
on grafted slides.
"""
import io, re, zipfile, json, subprocess, sys
from pathlib import Path

import pytest
from pptx import Presentation

WS = Path(__file__).parent.parent / "authoring" / "workspace"

def _pick_3_templates_from_first_deck():
    """Deterministic pick — alphabetically first deck's first 3 slides."""
    decks = sorted(p for p in (WS / "decks").iterdir() if p.is_dir())
    for d in decks:
        slides = sorted((d / "slides").glob("slide_*.yaml"))
        if len(slides) >= 3:
            import yaml
            return [yaml.safe_load(s.read_text())["id"] for s in slides[:3]]
    pytest.skip("need a deck with at least 3 slides in workspace")

def _all_rids_resolve(pptx_bytes):
    """Return list of (slide_name, rid) for broken references; empty list = OK."""
    broken = []
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for name in z.namelist():
            m = re.fullmatch(r"ppt/slides/(slide\d+)\.xml", name)
            if not m:
                continue
            slide_name = m.group(1)
            xml = z.read(name).decode("utf-8")
            used = set(re.findall(r'r:(?:embed|link|id)="(rId\d+)"', xml))
            rels_name = f"ppt/slides/_rels/{slide_name}.xml.rels"
            try:
                rels = z.read(rels_name).decode("utf-8")
                have = set(re.findall(r'Id="(rId\d+)"', rels))
            except KeyError:
                have = set()
            for rid in used - have:
                broken.append((slide_name, rid))
    return broken

def _compose_via_subprocess(plan_json: str) -> bytes:
    """Run the compose pipeline end-to-end; returns the .pptx bytes."""
    # Invoke via the local Flask app or direct reader.py call;
    # implementation detail.
    ...

def test_three_slide_compose_no_broken_rels():
    ids = _pick_3_templates_from_first_deck()
    plan = [{"template": tid, "slots": {}} for tid in ids]
    out = _compose_via_subprocess(json.dumps(plan))
    broken = _all_rids_resolve(out)
    assert not broken, f"broken rid references: {broken}"

def test_three_slide_compose_opens_in_python_pptx():
    ids = _pick_3_templates_from_first_deck()
    plan = [{"template": tid, "slots": {}} for tid in ids]
    out = _compose_via_subprocess(json.dumps(plan))
    prs = Presentation(io.BytesIO(out))
    assert len(prs.slides) == 3

@pytest.mark.phase3
def test_grafted_slides_have_explicit_bg():
    ids = _pick_3_templates_from_first_deck()
    plan = [{"template": tid, "slots": {}} for tid in ids]
    out = _compose_via_subprocess(json.dumps(plan))
    with zipfile.ZipFile(io.BytesIO(out)) as z:
        for n in ("slide2.xml", "slide3.xml"):
            xml = z.read(f"ppt/slides/{n}").decode("utf-8")
            assert "<p:bg" in xml, f"{n} has no explicit background"
```

The first two tests pass after Phase 1; the third only after Phase 3.

For PowerPoint-render verification (the "no repair dialog" criterion),
no programmatic test exists. Manual: open in PowerPoint on macOS or
Windows; ensure no warning dialog appears.

---

## Open decisions for the engineer

1. **Implement flatten at compose-time or at ingest-time?** The doc
   above puts it at compose-time (inside `_copy_slide_into`). The
   alternative is at ingest (`write_slide_fragment`), so all
   workspace fragments are pre-flattened. Compose-time is more
   flexible (re-ingesting isn't needed when the engine changes); ingest-time
   is faster at compose-run. **Recommend compose-time** for the
   first cut.

2. **Hyperlinks rel-import path**. python-pptx's `relate_to` may not
   handle external-target rels cleanly. May need a thin
   `_add_external_rel` helper. Test early — pick a template with a
   hyperlink in its slot text and verify it survives the round-trip.

3. **Idempotency on already-flattened slides**. The
   `_flatten_slide_background` helper should be idempotent (running
   twice produces the same XML). Add an assertion early.

4. **What about `prune_unreachable_parts` at ingest time?** If
   flatten is compose-time and successfully resolves all background
   inheritance, the ingest-side prune is fine as-is. If you discover
   the prune is dropping parts the flatten needs, fix the prune's
   closure walk to follow background-fill chains. Out of scope for
   the initial PRs but flag if you see it.

5. **Should we accept slow-compose for correctness?** Flatten adds
   per-slide work (theme XML parsing, deep clone, rel imports). For a
   5-slide deck, expect ~100ms extra. If compose latency becomes a
   product concern, cache per-deck theme parses across the compose
   run. Not premature optimization unless someone complains.

---

## Things explicitly out of scope

- Compose-mode (`{"compose": true, ...}`) implementation. That's
  Phase D / a separate piece of work.
- Cross-deck color/font remap correctness. The existing
  `_apply_scheme_remap` / `_apply_font_remap` calls stay in place;
  the rel-aware import works alongside them.
- Replacing python-pptx with a different library.
- Modifying `authoring/cli.py`'s ingest path. Bug 2 has an ingest-side
  manifestation but we're fixing it at compose-time.

---

## How to verify your fix locally

1. Run the repro script in this doc against current main. Confirm you
   see `slide2.xml missing: rIdN` (Bug 1 active).
2. Implement Phase 1.
3. Re-run repro. Confirm all slides report `missing:` (empty).
4. Open `/tmp/multi.pptx` in PowerPoint — no repair dialog.
5. Implement Phase 2; pick a template with a chart atom; verify it
   appears in the output.
6. Implement Phase 3; visually compare backgrounds against the
   original source deck.
7. Run `pytest tests/test_multi_slide_compose.py -v`.
8. Rebuild the app zip via `python3 authoring/cli.py package-app`.
9. Re-transfer to the other PC if the original tester wants to
   confirm against their failing deck.

Push to a feature branch (`fix/multi-slide-compose-rels`) and open a
PR with the regression test in the PR description.
