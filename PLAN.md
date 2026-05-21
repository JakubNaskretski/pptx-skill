# pptx-skill — design plan

A two-layer system for turning example presentations into a portable,
agent-readable template library.

## The split

**Authoring layer** — lives in this repo. Runs on your machine. Takes
example `.pptx` files, strips them into structured templates, and lets
you describe each visual asset/slide at your own pace. Emits a portable
zip artifact.

**Consumer skill** — the zip artifact. Distributed to any consuming
agent (Copilot, Claude, GPT, internal tools). Read-only, dependency-
light, vision-agnostic.

The contract between them is the zip's file layout. Authoring is free
to change however it needs; the consumer artifact stays stable.

---

## Why this shape

Prior takes (`LokiAgents/LokiTest` v1, v2) collapsed under their own
weight: 100+ atomic CLI methods, three storage tiers, FastAPI review
app, brand-asset promotion, dedup heuristics. The vision/description
problem never actually got solved — it was punted to "either a human
or some other agent." The describe step is the load-bearing piece, so
v3 treats it as the primary surface.

Locked design choices for v3:

| | Decision |
|---|---|
| Vision API | **None.** Skill is vision-agnostic. Descriptions arrive via the user (hand-edit YAML) or any external LLM the user has access to (paste image + bundled prompt). |
| Description workflow | **Resumable queue.** Ingest produces empty stubs; user works the queue over days/weeks. State persists in `workspace/`. |
| Authoring/consumer split | **Two layers, one zip contract.** Authoring complexity stays in this repo; agents only see the zip. |
| Consumer surface | **Three commands**: `list`, `get`, `compose`. No atomic shape ops. |
| Template granularity | **Whole slides only.** No individual-shape capture, no cross-deck dedup in v1. |
| Auto-promotion | Complete YAMLs auto-promote `pending → done` on `validate`. `locked: true` freezes a file. |

---

## Authoring layer

```
authoring/
  cli.py                       not yet implemented
  requirements.txt             python-pptx, pyyaml, click (TBD)
  prompts/
    describe_asset.md          paste-into-vision-LLM prompt
    describe_slide.md
  schemas/
    asset.yaml                 commented stub for new asset sidecars
    slide.yaml                 commented stub for new slide sidecars
  workspace/                   working state — gitignored
    decks/<stem>/
      original.pptx
      slides/slide_NN.pptx     extracted single-slide fragment
      slides/slide_NN.yaml     sidecar (stub → filled-in)
      slides/slide_NN.png      optional preview
    assets/
      <sha1>.<ext>             extracted image binary
      <sha1>.yaml              sidecar
  dist/                        gitignored
    skill.zip                  built consumer artifact
```

### Commands

| Command | Behavior |
|---|---|
| `ingest <deck.pptx>` | Strip a deck into `workspace/`. python-pptx + heuristic slot detection. Idempotent — existing YAMLs with `status: done` are not touched, structural fields refresh. |
| `status` | Counts pending/done; lists pending file paths. |
| `next [--kind asset\|slide] [--open]` | Print next pending item. With `--open`: launch `$EDITOR` on the YAML and reveal the asset in Finder. |
| `prompt [--kind asset\|slide]` | Print the bundled describe prompt to stdout for copy/paste into any vision-capable LLM. |
| `validate` | Schema-check every YAML; auto-promote complete ones to `status: done` unless `locked: true`. Report failures. |
| `preview` | Best-effort PNG thumbnails of slide fragments via LibreOffice. Silent skip if not installed. |
| `build [--allow-pending]` | Emit `dist/skill.zip`. Refuses if anything is `pending` unless `--allow-pending`. |

### Slot detection (ingest heuristics)

No AI. python-pptx exposes everything needed.

- `PP_PLACEHOLDER.TITLE` → `text` slot, `max_chars` = current text length × 1.5
- `PP_PLACEHOLDER.BODY` with bullets → `bullets` slot, `max_items` = current count
- `PP_PLACEHOLDER.PICTURE` or `Picture` shape > 20% slide area → `image` slot, `aspect` from geometry
- Everything else (logos, decorations, background fills) → frozen, not a slot

Each slot gets a stable id (`title`, `subtitle`, `body`, `hero`, …)
inferred from placeholder kind + position. Where two slots collide on
the same default id, suffix with index (`hero_left`, `hero_right`).

### Describe loop

1. `ingest big_deck.pptx` → produces N stubs marked `status: pending`.
2. `status` → e.g. `23 slides + 41 assets pending`.
3. `next` → tells you which file to describe + which YAML to fill.
4. Either:
   - Drag image + output of `prompt --kind asset` into Claude/ChatGPT/Gemini, paste returned YAML.
   - Hand-edit YAML using the inline schema comments as guide.
5. `validate` auto-promotes complete YAMLs to `done`.
6. Stop and resume tomorrow if needed — state persists in `workspace/`.
7. Eventually `status` shows 0 pending → `build`.

---

## Consumer skill (the zip)

```
pptx-skill/                    contents of skill.zip
  SKILL.md                     agent reads on load — small
  index.json                   flat list of templates + assets + tags
  templates/<id>/
    slide.pptx                 reusable slide fragment
    meta.yaml                  full description (intent, slots, suitable_for, …)
    preview.png                optional
  assets/<id>.<ext>
  assets/<id>.yaml
  reader.py                    ~150 LOC, the only code in the zip
  requirements.txt             python-pptx, pyyaml
```

### Agent-facing commands

| Command | Behavior |
|---|---|
| `python reader.py list [--filter feel=warm,kind=image]` | Return `index.json`, optionally filtered. JSON to stdout. |
| `python reader.py get <id>` | Return one item's `meta.yaml` + resolved file paths. JSON to stdout. |
| `python reader.py compose <plan.json> <out.pptx>` | Concatenate slide fragments, fill text + image slots, write the output deck. |

### Agent loop

1. Read `SKILL.md` on skill load.
2. `list --filter` to shortlist (e.g. `suitable_for=opener,feel=punchy`).
3. `get <id>` on the most promising to confirm slots.
4. Build a plan, call `compose`.

No vision required at consumption time. No API keys. No state.

---

## Build sequencing

Smallest end-to-end loop, no premature features:

1. **`ingest`** — python-pptx + heuristic slot detection + stub writer.
2. **`validate`** + auto-promotion — round-trip a hand-filled YAML.
3. **Describe prompts** — the real product of careful design; iterate
   on a few real assets before locking the schema.
4. **`status` + `next`** — resumability glue.
5. **`build`** — zip emitter + `index.json` builder.
6. **`reader.py`** — three commands on the consumer side.

After (1)–(6) there is a working end-to-end slice. Then:

7. `preview` (LibreOffice integration, optional).
8. Multi-deck workflow + cross-deck conventions.
9. Additional asset kinds beyond raster images, if needed.

---

## Deferred / not in v1

- Individual shape capture (callouts, banners, decorative shapes).
- SmartArt / OLE / embedded video as structured elements.
- Cross-deck dedup, "brand asset promotion".
- Cloud storage backends for assets (SharePoint, S3, HTTP).
- Brand-vocabulary token system (Primary/Secondary/Accent colors).
- Local review app — replaced by `next --open` + `$EDITOR`.
- Vision API integration — explicitly out of scope; descriptions arrive
  via the user or any external LLM they choose to use.
