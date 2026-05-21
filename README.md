# pptx-skill

A two-layer system for turning example presentations into a portable,
agent-readable template library.

- **Authoring layer** (`authoring/`) — tooling that ingests example
  `.pptx` files, strips them into structured slide templates with
  placeholder slots, and lets you describe each visual asset/slide at
  your own pace. Lives in this repo. Runs locally.
- **Consumer skill** (`consumer/`) — the portable artifact emitted by
  the authoring layer's `build` command. A small zip that any agent
  (Copilot, Claude, GPT, internal tools) can read to pick templates
  and compose decks. Read-only. Vision-agnostic. Three commands.

See [`PLAN.md`](PLAN.md) for the full design.

## Status

End-to-end v1 implemented: `ingest`, `status`, `next`, `prompt`,
`validate`, `preview`, `build` on the authoring side; `list`, `get`,
`compose` on the consumer side.

## Quickstart

```bash
# 1. Install authoring deps (one-time).
pip install -r authoring/requirements.txt

# 2. Strip a deck into the workspace.
python authoring/cli.py ingest path/to/your_deck.pptx

# 3. Describe slides and assets at your own pace.
python authoring/cli.py status
python authoring/cli.py next --open          # opens YAML in $EDITOR
python authoring/cli.py prompt --kind asset  # paste into a vision LLM

# 4. Validate. Complete sidecars auto-promote pending → done.
python authoring/cli.py validate

# 5. (Optional) Generate PNG previews if LibreOffice is on $PATH.
python authoring/cli.py preview

# 6. Build the portable consumer artifact.
python authoring/cli.py build
# → authoring/dist/skill.zip
```

The zip is the deliverable to consuming agents. They use
`reader.py list`, `get`, and `compose` against it — see
[`consumer/SKILL.md`](consumer/SKILL.md).
