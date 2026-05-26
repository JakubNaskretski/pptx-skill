# Asset selection — slim plan (POC)

Status: design, not implemented. Plan only; no code changes yet.

## Goal

Cut the asset-filter surface from **6 dimensions + 2 free-text fields** down
to **3 fields total** (`kind`, `tags`, `description`) plus mechanical
dimension metadata that the agent uses through helper methods, not as a
filter input.

This is a POC simplification — we trade some expressive power for
deterministic, low-drift picks the user can actually reason about.

## What changes — at a glance

### Kept (mechanical, never drifts)

| Field | Why |
|---|---|
| `kind` | Structural triage — required, validated against enum |
| `width`, `height`, `aspect` | Auto from binary at ingest; powers `check-asset-fit` |
| `colors_hex` | Auto from binary; used by recolor engine |
| `recolor_targets` | Auto for SVG; used by compose |
| `table` / `chart` / `shape` / `smartart` blocks | Auto for structured atoms |
| `id`, `sha1`, `sources`, `status` | Identity + workflow |
| `notes` | User-editable, harmless |

### New

| Field | Type | How populated |
|---|---|---|
| `tags` | list of enum, 0–4 entries | Vision LLM picks from `workspace/tag_vocab.yaml` |
| `description` | one short sentence (≤25 words) | Vision LLM; tiebreak only, never a filter |

### Dropped (subjective, drifts between describe runs)

`feel`, `composition`, `suitable_for`, `scope`, `colors` (word list),
`subject`, `depicts`, `interpretation`.

Existing `subject` / `depicts` content is **migrated into `description`**
(see migration section) so nothing user-authored is lost.

## Why no orientation/aspect filter dimension

You asked whether a strictly-technical companion tag to `kind` made
sense. After looking at it: mechanical dimensions are more useful
**as a helper method than as a categorical filter**. Reasons:

- Three buckets (`landscape` / `portrait` / `square`) lose information.
  A 16:9 hero slot and a 4:3 slot are both "landscape" but want
  different assets.
- An exact-fit / crop check is a one-line math operation we already
  promised in `REDESIGN.md` (`check-asset-fit`).
- Filtering by orientation forces the agent to make a choice; a fit
  check lets us answer "yes / no / yes with crop region" against the
  actual slot geometry.

So: store the raw numbers (`width`, `height`, `aspect`), expose them
through methods, don't expand the filter vocabulary.

```
find-asset --kind photo --tags people,office
  → returns shortlist of N assets, each with their dimensions

check-asset-fit asset_3e58c8f9 1920x1080
  → { fits: true, will_crop: { ... } }   (or fits: false + reason)
```

## Tag vocabulary — workspace-managed, user-editable

The closed enum lives in `workspace/tag_vocab.yaml`, **not** in
`authoring/schemas/vocab.yaml`. This is the one place where the user
can curate vocabulary without editing the source schemas.

### Seed list (POC, ~14 tags)

```yaml
# workspace/tag_vocab.yaml
tags:
  - people
  - office
  - laptop
  - device
  - screen
  - hands
  - document
  - chart
  - logo
  - abstract
  - outdoor
  - nature
  - city
  - workplace
```

These are deliberately broad. Domain-specific tags (`agriculture`,
`finance`, `healthcare`) get added by the user via UI as the library
grows, not seeded up front.

### Add a tag

CLI: `python3 authoring/cli.py tag-vocab add <tag>`
UI: button on a new "Tags" page in the sidebar.

Validation: lowercase, no spaces, dash-separated, must not already
exist.

### Remove a tag — forced remap

Removing a tag that's in use must answer "what happens to assets that
had it?". Two paths:

1. **Replace** with an existing tag, or
2. **Replace** with a new tag (which is also added to the vocab).

CLI: `python3 authoring/cli.py tag-vocab remove <tag> --replace-with <other>`
UI: removal dialog with two inputs (existing-tag dropdown OR new-tag
text field). Submit is disabled until one is chosen.

Unused tags can be removed without `--replace-with`.

### Validation

`python3 authoring/cli.py validate` checks that every tag on every
`<sha>.yaml` exists in `tag_vocab.yaml`. Drift here is a hard error,
because it means a tag was deleted without the remap step.

## New asset YAML shape

```yaml
# Identity (auto)
id: asset_3e58c8f9
sha1: 3e58c8f93e6f8bcd4db69df18cc98584a4191f83
sources:
  - deck: testPres3Mars
    slide: 6

# Type (mechanical at ingest, validated against enum)
kind: photo

# Mechanical visual (auto from binary)
width: 1920
height: 1080
aspect: 1.778            # stored as float; "16:9" derived for display
colors_hex:
  - '#49372A'
  - '#ADA48A'
  - '#755944'
recolor_targets: []       # vectors only

# Soft description (vision LLM, low-cardinality, idempotent)
tags:
  - people
  - nature
description: "Man standing among rows of small plants in a greenhouse-like enclosure"

# Kind-specific structural (auto; only present for structured atoms)
# table: { rows, cols, headers, sample_cells }
# chart: { type, series_count, categories_count }
# shape: { geometry, is_recolorable }
# smartart: { layout, nodes }

# Workflow
status: done
notes: ""
```

## Migration policy

Rule: **delete trash, preserve user content, fill blanks only.**

For each existing `workspace/assets/<sha>.yaml`:

### Delete unconditionally (the "trash")

`feel`, `composition`, `suitable_for`, `scope`, `colors`,
`interpretation`. These are gone — fields removed, not blanked.

### Preserve existing free-text by merging into `description`

- If `description` doesn't exist yet **and** `subject` is non-empty
  → `description = subject`.
- Else if `description` doesn't exist yet **and** `depicts` is
  non-empty → `description = depicts`.
- Else `description = ""` (only when both source fields are empty).

Then delete `subject` and `depicts` fields. The text content survives
in `description`; the field names go.

If `description` already has user content (manually set), **don't
overwrite**.

### Add new mechanical fields

`width`, `height`, `aspect` are computed from the binary at migration
time. Always added; cheap to recompute.

### Add `tags`

`tags: []` for every asset. Empty list. User curates per-asset via UI
or via a fresh describer pass on assets with `status: pending`. We
do not auto-populate from old `suitable_for` because the semantics
don't map cleanly.

### Status

Don't touch. `pending` stays `pending`; `done` stays `done`. New
assets that need a tag pass can be filtered on `tags == []` from the UI.

## File-by-file changes

### `authoring/schemas/vocab.yaml`

Strip `asset.feel`, `asset.composition`, `asset.suitable_for`,
`asset.scope_prefixes`. Keep `asset.kind`. (Slide vocab is untouched.)

### `authoring/schemas/asset.yaml`

Replace the descriptive-fields block with the new shape. Drop sections
for the removed fields. Add `tags`, `description`, `width`, `height`,
`aspect`. Update the inline docs at the top.

### `authoring/cli.py`

- Remove `ASSET_FEEL_ENUM`, `ASSET_COMPOSITION_ENUM`, `ASSET_SUITABLE_ENUM`,
  `ASSET_SCOPE_PREFIXES` (around line 1557).
- Add `TAG_VOCAB_PATH = WORKSPACE / "tag_vocab.yaml"` + load helper.
- Update `validate_asset()` to validate `tags[]` against tag vocab,
  drop checks on removed fields.
- Add `tag-vocab list / add / remove` subcommands.
- Update `_add_asset_to_workspace` (line 1945) to compute width/height/
  aspect via Pillow at ingest and write them into the sidecar template.
- Update `_asset_sidecar_template` to emit the new shape.
- Update `build-v5` asset record (line 2831) to ship the new fields
  (`tags`, `description`, `width`, `height`, `aspect`, `colors_hex`)
  in `index.json`. Drop the old structural-vocab block.
- New subcommand: `cli.py migrate-asset-yaml` — runs the migration
  policy above across all `workspace/assets/*.yaml`. Dry-run by default.

### `authoring/prompts/describe_asset.md`

Rewrite for the slim shape. The describer now picks `kind` (already
auto-seeded), `tags` (1–4 from the vocab list shown to it inline),
and writes one `description` sentence. Drop all guidance about
`feel`, `composition`, `subject`-vs-`depicts`, `scope` prefixes.
The prompt becomes ~⅓ its current length.

### `authoring/app.py`

- New `/tags` route + page: list tags, add, remove-with-remap dialog.
- Asset cards on the home page: show new fields, hide removed ones.
- `POST /api/asset/add` is unchanged externally (still returns the
  same shape) but internally emits the new sidecar template.

### `consumer/reader.py`

- `find-asset` (line 2283): drop `--feel`, `--composition`,
  `--suitable-for`, `--scope`, `--colors`. Add `--tags` (list).
  Broadening sequence becomes: drop `--tags` → fall back to placeholder.
- New `check-asset-fit` (already planned in REDESIGN.md): given an
  `asset_id` + target slot dimensions, returns
  `{fits, will_resize_to, will_crop}`. Pure math against
  `width`/`height`/`aspect`.
- `compose-v5` placeholder sentinel: unchanged.

### `consumer/SKILL_v5.md`

Rewrite the "Picking images" section. Algorithm becomes:

1. `find-asset --kind <k> --tags <t1>,<t2>`.
2. If empty: retry without `--tags` (one broadening step, that's it).
3. Among the shortlist, optionally call `check-asset-fit` against the
   slot to filter out aspect-incompatible candidates.
4. Pick by `description` text fit to slide topic.
5. Required-and-empty fallback: `POST /api/asset/add` or
   `"placeholder"` (both unchanged).

### `consumer/index.json` schema

Asset records carry the new fields. Old fields (`feel`, `composition`,
`suitable_for`, `scope`, `subject`, `depicts`) are not emitted.

## Implementation order (suggested)

1. **Schema + vocab** — `vocab.yaml`, `asset.yaml`, new tag vocab file.
2. **Migration script** — `cli.py migrate-asset-yaml`. Tested
   thoroughly on `TestData/` copies before touching real workspace.
3. **Ingest + add-asset** — wire new shape into `_add_asset_to_workspace`
   and `_asset_sidecar_template`.
4. **Describer prompt** — rewrite. Validate against one real asset.
5. **Validate + tag-vocab CLI** — `validate`, `tag-vocab` subcommands.
6. **`build-v5`** — emit new shape into `index.json`. Rebuild bundle.
7. **`find-asset` + `check-asset-fit`** — consumer side.
8. **`SKILL_v5.md`** — rewrite the picking section.
9. **UI** — `/tags` page, asset card updates.

Each step is independently testable. Steps 1–6 can ship together;
7–9 can follow on a second pass.

## What I'm not changing

- v5 compose path / placeholder sentinel.
- Slide vocab (`slide.feel`, `slide.suitable_for`).
- Skeleton / theme / decoration handling.
- `add-asset` external API (only the internal sidecar template
  changes).
- `remove-deck` behavior or the `external: manual-upload` sources
  variant — both stay.

## Open questions to resolve before coding

- **Field name for the technical tag.** `aspect` (float) vs.
  `aspect_ratio` (string like `"16:9"`). I lean float — easier
  math, no parsing — with a display helper that formats it.
- **Tag count cap.** Plan says 0–4 per asset. If you'd rather 1–3
  to keep things tighter, that's one validator constant.
- **Where the tag-vocab UI lives.** New `/tags` page vs. inline on
  the home sidebar. I lean separate page — keeps the home view
  about assets, not metadata management.
- **Re-describe button.** Should the UI have a "re-describe this
  asset with new prompt" button to easily fill `tags` on legacy
  assets? Or rely on the user editing the sidecar directly? I'd say
  yes to a button — without it, migrating the existing library is
  a hand-edit slog.
