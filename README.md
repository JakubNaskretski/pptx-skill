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

Phase 0 — scaffolding only. Schemas, prompts, and the design plan are
checked in. No working CLI code yet. Implementation is sequenced in
PLAN.md → "Build sequencing".
