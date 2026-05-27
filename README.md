# pptx-skill

Turn example PowerPoint decks into a portable, agent-readable library of
**slide skeletons** and **visual assets**, then have an LLM compose new
decks against that library — no vision needed at consume time.

Two layers:

- **Authoring layer** (`authoring/`) — a local Flask app + CLI that
  ingests `.pptx` files, extracts their slide structures (skeletons)
  and visual atoms (photos, logos, icons, tables, charts), and lets
  you review/describe each one through a browser UI.
- **Consumer skill** (`consumer/`) — a portable `.zip` artifact emitted
  by the authoring layer's `build-v5` command. Any agent (Claude,
  ChatGPT, Copilot) can read it to pick a skeleton, fill its slots,
  and render a `.pptx`. Read-only, vision-agnostic, runs against a
  small `reader.py`.

## Quickstart

```bash
# 1. Install authoring deps (one-time).
pip install -r authoring/requirements.txt

# 2. Strip a deck into the workspace.
python3 authoring/cli.py ingest path/to/your_deck.pptx

# 3. Launch the local web app to review skeletons + describe assets.
python3 authoring/app.py
# → http://localhost:5050 opens automatically

# 4. (Optional) Render skeleton preview PNGs so the review UI has
#    thumbnails. Tries PowerPoint (Windows), LibreOffice, then macOS
#    Quick Look. Skipped silently if no renderer is available.
python3 authoring/cli.py preview

# 5. Build the portable consumer artifact.
python3 authoring/cli.py build-v5
# → authoring/dist/skill-v5.zip
```

The zip is the deliverable to consuming agents — see
[`consumer/SKILL_v5.md`](consumer/SKILL_v5.md) for the contract.

---

## Architecture in one screen

```
┌─ Your PowerPoint decks (.pptx) ──────────────────────────────────────┐
│                                                                      │
│                              ingest                                  │
│                                ▼                                     │
└─ authoring/workspace/ ───────────────────────────────────────────────┘
  │
  ├── themes/<deck>/              ← extracted from each ingested deck:
  │   ├── theme.yaml                palette (primary/accent/text/bg) +
  │   ├── master.pptx               theme fonts + the deck's slide master
  │   └── preview.png               so v5 compose can re-host any
  │                                 skeleton on this look.
  │
  ├── skeletons/<deck>_<NN>/      ← one per ingested slide:
  │   ├── skeleton.yaml             structural layout — slot list with
  │   ├── preview.png               kind/geometry/style/constraints,
  │   └── background.png            categories (opening/content/data/
  │                                 closing), source_deck pointer.
  │                                 Skeletons are reviewed/approved
  │                                 via the web UI.
  │
  ├── assets/<sha1>.{yaml,bin}    ← one per unique image/icon/atom
  │                                 found across all ingested decks
  │                                 (deduped by content hash):
  │                                   slim sidecar (kind, tags,
  │                                   description, width/height/aspect,
  │                                   colors_hex) + the binary itself
  │                                   (png/jpg/svg/xml).
  │
  ├── tag_vocab.yaml              ← workspace-editable closed tag list
  │                                 used by every asset's `tags` field
  │                                 (managed via `cli tag-vocab`)
  │
  └── decks/<deck>/               ← only `slide_NN.pptx` fragments
                                    kept here — needed by the preview
                                    renderer (qlmanage / LibreOffice).
                                    No v4 slide YAMLs.

           │
           │   build-v5  OR  /compose page in the web app
           ▼

┌─ The brief bundle — what ships to the agent (text-only) ─────────────┐
│                                                                      │
│   SKILL.md          ← agent contract (`SKILL_v5.md` content)         │
│   reader.py         ← the agent calls this for find-asset,           │
│                       check-asset-fit, match-skeletons               │
│   tag_vocab.yaml    ← live tag list                                  │
│   brand.md          ← (optional) per-org palette / voice / taboos    │
│   brief.md          ← user's deck brief (typed in /compose)          │
│   index.json        ← v5 summary catalog:                            │
│                       {version: 5, themes, skeletons, assets,        │
│                        tag_vocab}                                    │
│   themes/<id>/theme.yaml      ← palette + fonts only                 │
│   skeletons/<id>/skeleton.yaml ← structural slot layout              │
│   assets/<id>.yaml            ← slim sidecar (kind, tags,            │
│                                  description, dimensions, colors)    │
│   user_assets/      ← (optional) user-attached images for THIS       │
│                       request — low-res previews in the zip,         │
│                       full-res spliced back in at compose time       │
└──────────────────────────────────────────────────────────────────────┘

The bundle is **text only**. No KB asset binaries, no rendered slide
previews, no master.pptx — the agent picks IDs by reading YAML; the
actual binaries live on the authoring machine and get spliced in
server-side at compose-run time. Typical bundle is ~100 KB for a
30-skeleton workspace.
```

### How the pieces relate

- **A skeleton** is the structural skeleton of one slide: where its
  text slots / image slots / tables / charts live, with constraints
  (max chars, aspect ratio, etc.). It does NOT include the original
  asset bytes — those live in `assets/` and are referenced by id at
  compose time. Categories (`opening` / `content` / `data` / `closing`)
  let the filter UI narrow what the agent sees.
- **An asset** is a single visual atom: a photo, an icon, a logo, a
  table fragment, a chart fragment. Each one has a slim record
  (kind, tags, description, dimensions, color palette) so the agent
  can pick a fitting asset for an image slot without ever seeing the
  pixels.
- **A theme** captures the source deck's palette + fonts + slide
  master. Any skeleton can be re-hosted on any theme at compose time —
  that's how cross-deck composition works.
- **The brief bundle** is what `/compose` ships to your LLM: SKILL.md
  + reader.py + brand.md + brief.md + filtered skeletons + matching
  themes + slim asset catalog + tag vocab. The LLM reads, writes a
  v5 plan (`[{skeleton_id, slots: {...}}, ...]`), pastes it back into
  `/compose`, and the same workspace re-stages a v5 bundle for
  `reader.py compose-v5` to render the `.pptx`.

### Web UI surface

Local Flask app on `http://localhost:5050`. Two pages.

- **`/` — review.** One tabbed page covering everything you ingest.
  Tab switcher in the sidebar:
    - **Skeletons** — every ingested slide grouped by deck. Checkboxes
      for bulk approve/reject; preview pane shows the slide with
      colored slot overlays; right panel lets you override slot
      kind / role, promote unmapped shapes, set overlap decisions.
      Auto-advances to the next pending after a single approve/reject.
    - **Assets** — every deduped image / icon / atom from every deck.
      Sidebar lists them by id with a Hide-done filter; preview pane
      shows the binary; right panel is the describe form
      (kind / tags / description / notes). Save promotes pending → done
      and jumps to the next pending automatically. `+ Add asset`
      uploads a single image without re-ingesting a deck.
  `/v5` redirects here for old bookmarks.
- **`/compose` — brief builder.** Left rail filters the library by
  skeleton `categories` and asset `kind` / `tags`. Middle pane is
  your brief text + optional user-attached assets. Bottom emits a
  zip bundle (download) or flat-text view (copy/paste). Plan from
  the LLM goes into the third pane and runs through `reader.py
  compose-v5` to produce the final `.pptx`.

### CLI surface

```bash
# Ingest + workspace lifecycle
python3 authoring/cli.py ingest <deck.pptx>           # extract skeletons + assets
python3 authoring/cli.py remove-deck <deck-stem>      # purge a deck's contribution
python3 authoring/cli.py preview                      # render skeleton thumbnails
python3 authoring/cli.py status                       # pending/done/locked counts

# Asset tagging
python3 authoring/cli.py tag-vocab list
python3 authoring/cli.py tag-vocab add <tag> --description "..."
python3 authoring/cli.py tag-vocab remove <tag> --replace-with <other>
python3 authoring/cli.py redescribe --all             # re-queue all assets for re-tag

# One-off helpers
python3 authoring/cli.py add-asset <file> --kind photo
python3 authoring/cli.py next --open                  # opens next pending sidecar
python3 authoring/cli.py validate                     # auto-promote complete ones

# Bundle
python3 authoring/cli.py build-v5                     # → dist/skill-v5.zip
python3 authoring/cli.py package-app                  # → dist/pptx-skill-app.zip
```

### Assumptions baked into the system

- **One workspace per user.** All ingested decks share the same
  `assets/` pool — duplicates dedupe by SHA1 of the binary.
- **Tags are a closed vocabulary.** The agent picks from
  `tag_vocab.yaml`; if a needed tag isn't there, the user adds it via
  CLI or accepts `tags: []`. The describer prompt inlines the live
  list every call so the LLM never invents tags.
- **Descriptions are short and visual.** One neutral sentence under
  30 words, what's literally pictured — not rhetorical meaning. Code
  enforces `<=30`; prompt recommends `<=25`.
- **Skeletons are immutable once approved.** Editing slot kinds /
  roles is supported (and stamped `user_edited`) but the structural
  layout comes from the source slide; we don't manually build
  skeletons from scratch.
- **Compose is v5-only.** Plans must be v5-shaped
  (`{"skeleton_id": ..., "slots": {...}}`); the v4 path
  (`{"template": ..., "slots": ...}`) was removed.
- **Theme picking at compose.** The first skeleton's `source_deck`
  determines the theme unless a plan entry overrides with
  `"theme": "<id>"`. If that deck has no `workspace/themes/<deck>/`
  registered, compose-run fails — the bundle can't carry a theme it
  doesn't have.

### What it doesn't do

- No chart-data rebinding yet (charts ship as static images / SmartArt
  grafts that the engine doesn't recolor on the fly).
- No live cloud sync — everything is on disk under
  `authoring/workspace/`.
- No multi-user / no auth — the Flask app binds localhost.
- No vision at consume time. The agent only sees `description` and
  `tags`; if those are wrong, picks will be wrong.

See [`CHANGELOG.md`](CHANGELOG.md) for the version-by-version diff and
[`consumer/SKILL_v5.md`](consumer/SKILL_v5.md) for the agent contract.
