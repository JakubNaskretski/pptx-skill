# Changelog

## Unreleased — v5-only: brief alignment, skeleton bulk review, v4 purge

Two PRs combined: [#9](https://github.com/JakubNaskretski/pptx-skill/pull/9)
(v5 brief + skeleton multi-select) and
[#10](https://github.com/JakubNaskretski/pptx-skill/pull/10) (v4
deletion). Net result: the entire system is v5-native; v4 commands,
v4 SKILL.md, v4 helpers, v4 slide YAMLs, and v4 compose-run paths
are gone.

### Why

After the slim-asset and structural-skeleton work landed, the v4
slide path (`decks/<deck>/slides/slide_NN.yaml` + the v4
`templates`/`assets` index shape) became a parallel dead universe.
The `/compose` brief page was still emitting that v4 catalog — so
the agent never saw any skeletons even though the user reviewed +
approved them via `/v5`. PR #9 closes that gap; PR #10 cleans up
what's left behind.

### `/compose` brief is now v5-native (PR #9)

- **`/api/compose/bundle`** and **`/api/compose/text`** ship the v5
  layout: `SKILL_v5.md` as `SKILL.md`, `reader.py`,
  `skeletons/<id>/{skeleton.yaml,preview.png,background.png}`,
  `themes/<id>/{theme.yaml,master.pptx,preview.png}` (trimmed to
  decks-in-bundle), `assets/<id>.{yaml,bin}`, `tag_vocab.yaml`, plus
  `brand.md` / `brief.md` / user-attached low-res previews. The
  `index.json` is `{version: 5, themes, skeletons, assets, tag_vocab}`.
- **`brief.md` slimmed** — the long v4 four-pass preamble + helper
  docs + slot-polymorphism notes were dropped; `SKILL_v5.md` is now
  the canonical contract.
- **Compose page filter UI** renames `Templates` → `Skeletons`;
  filter dimensions become skeleton `categories` and asset `kind` /
  `tags`. `KB summary` and match-count text updated.
- **Same workspace** drives both review (`/v5`) and brief (`/compose`).
  Skeletons approved in `/v5` are immediately picked up by the brief
  unless their status flips to `rejected`.

### `/v5` skeleton review (PR #9)

- **Multi-select** — every sidebar item gets a checkbox. Selecting
  any reveals a sticky bulk-action bar (Mark done / Reject / Clear).
  ≥5 selected prompts a confirm.
- **Auto-advance** — after a single Mark-done or Reject (or
  overlap-Reject) the panel jumps to the next pending skeleton in
  list order. Corrective actions (Back-to-pending, Restore) stay put
  so the user can keep editing.
- Selection state survives sidebar reloads; stale ids are pruned.

### v4 deletion (PR #10)

Net: 14 files, +96 / -2,730 lines. Nothing left calls any of these.

- **`cli.py`** — removed `build` command, `build_index`,
  `_template_index_entry`, `_asset_index_entry`, `iter_slide_yamls`
  + every caller, `validate_slide` + `SLIDE_*` enums, `infer_layout`,
  `write_slide_yaml_stub`. `cli next` / `prompt` / `status` no longer
  take `--kind`.
- **`_ingest_pptx`** stops writing `decks/<deck>/slides/slide_NN.yaml`
  and the v4 `decks/<deck>/theme.yaml`. Per-slide `.pptx` fragments
  are still written — the preview renderer still needs them.
- **`app.py`** — removed `_stage_compose_bundle`,
  `_plan_looks_like_v5`, the else-branch of `/api/compose/run`,
  `_bulk_instructions_slide_with_assets`, `_assets_for_slide`,
  `_ensure_slide_png`, slide branches in `/api/batch/create` and
  `/api/batch/apply`, `/preview` slide branch, `SLIDE_DESCRIPTIVE`.
- **`schemas/vocab.yaml`** — dropped the `slide:` block.
- **`prompts/describe_slide.md`** — deleted.
- **`consumer/SKILL.md`** + **`consumer/helpers/`** — deleted.
  `SKILL_v5.md` is the canonical contract.
- **`consumer/reader.py`** — removed `cmd_list` / `cmd_get` /
  `cmd_compose` + argparse wiring. v4 path helpers
  (`template_dir`, `asset_path`, `asset_meta_path`, `parse_filter`,
  `matches_filter`) and the v4 compose engine internals (`_fill_slot`
  family, compose-mode shape placer) are still in the file but
  unreferenced from the CLI — a deeper consumer-side cleanup is a
  separate workstream.

### Migration from pre-#9 workspace

If you've been using the local app pre-#9, nothing on disk needs to
change. The first compose bundle you generate after pulling will be
v5-shaped automatically. If you're using the `cli build` zip in
production: it's gone — switch to `cli build-v5` which produces
`dist/skill-v5.zip` with the v5 contract.

Stale `decks/<deck>/slides/*.yaml` files left over from old ingests
sit on disk but are no longer read by anything. Safe to `rm -rf`;
no auto-purge tool ships in this PR.

### Known gap

Two of the four decks in the test workspace
(`standard_templates`, `title_and_breaks`) don't have
`workspace/themes/<deck>/theme.yaml` registered. The brief bundle
correctly trims themes to decks-in-bundle, but if the agent picks a
skeleton from one of those decks, compose-run will fail. Workaround:
re-ingest those decks, or restrict the brief filter to decks with
themes. A proper "missing-theme" warning surface is on the TODO
list.

## Earlier unreleased — slim asset schema + workspace tag vocab

### Why

The describe-time soft fields (`feel`, `composition`, `suitable_for`,
`scope`, `colors`-as-words, `subject`, `depicts`, `interpretation`)
drifted across vision-LLM passes and made `find-asset` queries hard
to reason about. Six filter dimensions for ~40 assets was overkill.

This pass cuts the agent-facing surface to **kind + tags +
description**, moves all visual dimension math onto mechanical fields
(`width`, `height`, `aspect`) the engine reads through
`check-asset-fit`, and lets the user curate the tag list at runtime
without editing source schemas.

### Schema (`authoring/schemas/`)

- **`asset.yaml`** rewritten. Dropped: `subject`, `depicts`, `feel`,
  `composition`, `colors` (word list), `scope`, `suitable_for`,
  `interpretation`. Added: `tags[]`, `description`, `width`,
  `height`, `aspect`. Kept: `kind`, `colors_hex`, `recolor_targets`,
  kind-specific blocks (`table` / `chart` / `shape` / `smartart`),
  `id` / `sha1` / `sources`, `status`, `notes`.
- **`vocab.yaml`** asset section trimmed to `kind` only. Slide vocab
  is unchanged.
- **`workspace/tag_vocab.yaml`** — new file, ships with a 14-tag seed
  list (`people`, `office`, `laptop`, `device`, `screen`, `hands`,
  `document`, `chart`, `logo`, `abstract`, `outdoor`, `nature`,
  `city`, `workplace`). User-editable; the validator enforces that
  every tag on every sidecar exists in this file.

### Authoring (`authoring/`)

- **`cli.py migrate-asset-yaml`** — new command. Strips dropped
  fields, folds `subject` (or `depicts` if `subject` is empty) into
  `description` only when `description` is empty, seeds `tags: []`,
  fills `width`/`height`/`aspect` from the binary. Idempotent; default
  is `--dry-run`; `--apply` writes a backup to
  `workspace/_migration_backup/<timestamp>/` before mutating.
- **`cli.py tag-vocab`** — new command group. `list` prints the
  current vocab; `add <tag>` appends a new tag; `remove <tag>` removes
  one and rewrites any sidecar using it. Removal of an in-use tag
  requires `--replace-with <other>` so no asset is left orphaned.
- **`cli.py redescribe <asset_id>`** / `--all` — new command. Marks
  an asset (or every asset) pending and clears `tags`. **Preserves
  `description`** so existing prose survives the round-trip. Use to
  feed the slim-shape describer over your existing library without
  re-ingesting.
- **`cli.py validate`** — rewired. Asset checks now: `kind` in
  enum, `tags` is a non-empty list of ≤4 entries each present in
  `tag_vocab.yaml`, `description` non-empty and ≤30 words
  (recommended ≤25).
- **`cli.py add-asset`** — unchanged externally. Internally writes
  the new slim sidecar shape, with `width`/`height`/`aspect` computed
  from the binary at ingest time.
- **`cli.py build-v5`** — asset records in `index.json` now ship the
  new shape; the bundle also carries a top-level `tag_vocab` array so
  the consuming agent can read the live vocabulary without poking
  workspace state.
- **`prompts/describe_asset.md`** — rewritten for the slim shape.
  Inlines the current tag vocabulary so the vision LLM has the closed
  list in front of it. The prompt is ~⅓ its previous length.

### Consumer (`consumer/`)

- **`reader.py find-asset`** — argument surface cut from six
  optional filters to one: `--kind` (required) and `--tags`
  (repeatable; AND-matched). Broadening reduces to a single step:
  drop `--tags` and retry. Response payload now ships
  `description`, `tags`, mechanical dimensions, and `colors_hex`
  per match, plus the shared `tag_vocab` for reference.
- **`reader.py check-asset-fit`** — unchanged externally; now reads
  the mechanical `width`/`height` straight off the sidecar, instead
  of relying on optional `dimensions` nested under the legacy shape.
- **`SKILL_v5.md`** — "Picking images" rewritten for the new
  algorithm. Bundle-layout footer updated.

### Migration

For an existing workspace with described assets:

```bash
python3 authoring/cli.py migrate-asset-yaml         # dry-run first
python3 authoring/cli.py migrate-asset-yaml --apply # writes; backs up
python3 authoring/cli.py validate                   # confirms slim shape
python3 authoring/cli.py build-v5                   # rebuild bundle
```

Migrated sidecars land with `tags: []` and `status: pending` — the
old `feel/composition/scope/...` fields don't map cleanly onto the
new closed `tag_vocab`, so re-curation is intentional. `validate`
will flag them as incomplete until a re-tagging pass (manual or via
`redescribe` + the describer prompt).

To re-describe existing assets against the new tag list without
losing the migrated `description`:

```bash
python3 authoring/cli.py redescribe --all
python3 authoring/cli.py next --kind asset --open   # walk the queue
# or use the web app's describe page
```

The migration never overwrites a non-empty `description`. Old
`subject` content is folded in only when `description` is empty.

## Earlier unreleased — idempotent asset selection + single-asset ingest

Branch: `claude/pptx-pr-review-57WHk` (rebased onto current `main`,
which already includes `cli.py remove-deck` — see "Interaction with
remove-deck" below).

### Why

Two related improvements driven by composition-time pain:

1. **Image selection was vibes-based.** The agent read the whole asset
   library out of `index.json` and picked an `asset_id` by free-text
   `subject` / `depicts`. Those fields drift between describe passes, so
   identical briefs gave different picks across runs. Compose was not
   idempotent on the asset axis.
2. **Adding one image required a full deck.** The only path into the
   asset library was `cli.py ingest <deck.pptx>` — you couldn't drop in
   a single PNG / SVG without bolting it into a pretend `.pptx` first.

### Consumer-side changes (`consumer/`)

- **`reader.py find-asset`** — new subcommand. Filters `index.json`'s
  assets by structural tags only (`--kind` required; `--feel`,
  `--composition`, `--suitable-for`, `--scope`, `--colors` optional).
  Returns a deterministic shortlist sorted by `id`, capped by `--limit`
  (default 5). Free-text fields (`subject`, `depicts`) are included in
  each match for tiebreaks but are NEVER part of the filter.
  - On empty result, the response includes a `suggestion` naming the
    most-specific constraint to drop first (broaden order:
    `--colors` → `--composition` → `--scope` → `--suitable-for` → `--feel`).

- **`reader.py compose-v5`** — placeholder sentinel. Passing
  `"placeholder"` (or `{"placeholder": true, "label": "..."}`) as an
  image-slot value now renders a dashed light-grey rectangle with the
  slot id as a label, and emits an `image_placeholder` warning to the
  sidecar so the user knows which slots still need a real asset.

- **`SKILL_v5.md`** — new *"Picking images"* section enforcing the
  algorithm:

  1. `find-asset` first with `kind` + deck-`feel` + slot `suitable_for`.
  2. If empty: drop the constraint from `suggestion`, retry.
  3. Pick from the shortlist by `subject` / `depicts`.
  4. Required-and-empty fallbacks: stage via `POST /api/asset/add`, or
     use the `"placeholder"` sentinel.

  The slot-kinds table now documents `"placeholder"` as a valid image
  value.

### Authoring-side changes (`authoring/`)

- **`cli.py add-asset <path>`** — new command. Takes one
  `.png|.jpg|.jpeg|.webp|.gif|.svg|.xml`, hashes the contents, copies
  the binary into `workspace/assets/<sha1>.<ext>`, writes a sidecar
  YAML with empty descriptive fields and `status: pending`. Auto-seeds
  `kind` from the extension (raster → `photo`, SVG → `vector`, XML →
  blank for hand-editing). For rasters, runs the existing PIL dominant-
  colour extractor; for SVGs, runs the existing fill/stroke parser to
  seed `colors_hex` and `recolor_targets`.
  - Optional `--kind` override (validated against the asset-vocab enum).
  - Idempotent: re-adding the same file is a no-op.

- **`POST /api/asset/add`** — new endpoint (multipart form field
  `file`, optional `kind`). Calls `cli_mod._add_asset_to_workspace`
  under the hood. Returns
  `{"asset_id", "sha1", "yaml_path", "binary_path", "kind"}`.

- **UI** — new `+ Add asset` button in the home-page sidebar, paired
  with the existing `+ Ingest .pptx` button. Same UX pattern: hidden
  file input + status message + sidebar refresh on success.

- **`build-v5`** — asset records in `index.json` now carry the full
  structural-vocab block (`feel`, `composition`, `suitable_for`,
  `scope`, `colors`, `colors_hex`) plus `depicts`. Previously only
  `id`, `kind`, `subject` rode along, which made consumer-side
  `find-asset` impossible. **Re-run `build-v5` to refresh existing
  bundles** — old bundles will still parse but `find-asset` will treat
  every asset as un-tagged.

- **`_prune_dead_sources`** — now keeps entries that lack a `deck` key
  (external / manual-upload sources). Previously the helper dropped any
  source whose `deck` wasn't in the workspace, which would have erased
  manually-added assets' provenance on the next re-ingest.

### Interaction with `remove-deck`

`cli.py remove-deck <stem>` (already on `main`) calls
`_prune_dead_sources` on every asset that lists the removed deck as a
source. With this change, an asset that was *also* manually uploaded —
whose `sources` carry both a `{deck, slide}` entry and an
`{external: manual-upload, …}` entry — keeps the external entry when
its parent deck is removed. The binary stays in `workspace/assets/`
and remains visible to `find-asset`. Pure deck-derived assets behave
exactly as before.

### Schema

- **`sources[]`** in `<sha1>.yaml` now accepts a second entry shape for
  non-deck origins:

  ```yaml
  sources:
    - external: manual-upload
      filename: photo.jpg
      at: 2026-05-26T13:42:11
  ```

  Deck-derived `{deck, slide}` entries are unchanged.

### Migration notes

- Existing un-described assets remain untagged — they'll be invisible
  to `find-asset` filters that specify `feel` / `suitable_for` / etc.
  until they're described. They still match `find-asset --kind <k>`
  with no other filters.
- Run `cli.py build-v5` to regenerate `dist/skill-v5.zip` with the
  richer asset records before deploying the new bundle.
- No breaking changes to v4 commands, plan shapes, or YAML schemas
  beyond the additive `sources[]` variant above.
