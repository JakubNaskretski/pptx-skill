# Findings — agent prompt-bundle improvements

Status: pre-implementation notes. Scope: the artifact produced by the
**Download bundle (.zip)** button on the Compose page —
[authoring/app.py:2476-2496](../authoring/app.py#L2476-L2496) →
`POST /api/compose/bundle` →
[`_build_prompt_bundle_zip()`](../authoring/app.py#L1058-L1086).

The bundle is what an LLM agent (Opus 4.6 / GPT-5.x / etc.) reads to
produce a `plan.json` array that the user pastes back into the UI for
`reader.py compose`. Two work areas:

1. **Bundle/index shape** — what data the agent sees, and where the
   indexing makes its job harder than it needs to be.
2. **In-bundle helper scripts** — small Python utilities shipped *in*
   the bundle so an agent with code execution can answer common
   catalog questions correctly instead of writing ad-hoc scripts that
   misread the schema.

---

## 1. Bundle / index shape

### What the bundle contains today

```
prompt-bundle-<ts>.zip
├── SKILL.md                    consumer/SKILL.md — agent instructions
├── brand.md                    optional — only when non-empty
├── index.json                  catalog (cli.build_index)
├── brief.md                    user brief + restated output rules
└── decks/<deck>/theme.yaml     only for decks present in filtered KB
```

Verified against the sample at `TestData/prompt-bundle-20260523-232710`
(local-only, git-ignored). 73 KB index.json over 3 decks.

Per-entry shape from
[`_template_index_entry` / `_asset_index_entry` at cli.py:2110-2154](../authoring/cli.py#L2110-L2154).

### Gaps worth fixing

#### G1. `sources.deck` is dropped from index entries

Sidecars carry `sources: [{deck, slide}]` but the index helpers don't
forward it. The agent therefore **cannot tell which deck a template
or asset came from**, yet the bundle ships `decks/<deck>/theme.yaml`
per deck. Adding a `"deck": "<stem>"` field to both index helpers is
~2 lines and unblocks same-deck consistency reasoning, theme lookups,
and cross-deck disambiguation when ids happen to be similar.

**Files:**
[cli.py:2110-2127](../authoring/cli.py#L2110-L2127),
[cli.py:2130-2154](../authoring/cli.py#L2130-L2154).

#### G2. No previews in the bundle

`cli build` emits `templates/<id>/preview.png`
([cli.py:2262-2264](../authoring/cli.py#L2262-L2264)), the compose
prompt-bundle does not. For multimodal agents a 30–60 KB JPEG is
worth more than the `layout` string + `inventory[]` block combined.
Gate behind an "Include previews" checkbox in the Compose UI so
text-only flows keep the bundle small.

**Files:**
[app.py:1058-1086](../authoring/app.py#L1058-L1086), Compose step 2
panel in [app.py:2311-2331](../authoring/app.py#L2311-L2331).

#### G3. `inventory[]` references atoms by id only

Today each template `inventory[]` entry is
`{atom: asset_X, kind, x, y, w, h, region}`. To learn what
`asset_X` actually is, the agent has to cross-reference the assets
array. Inlining 2–3 fields (`subject` snippet, `feel`, `colors_hex`)
into each inventory entry cuts the cross-lookups during template
scoring. Cost: small JSON size bump per template.

**File:** [cli.py:2110-2127](../authoring/cli.py#L2110-L2127).

#### G4. `brief.md` duplicates large chunks of `SKILL.md`

[`_format_brief` at app.py:991-1055](../authoring/app.py#L991-L1055)
re-states v4 capabilities, slot polymorphism, the bullets warning,
compose-mode acceptance — all of which is in `SKILL.md` too. When
SKILL.md changes the brief copy rots silently. Trim brief.md to:
brief text + brand reminder + one-line "output JSON array only,
read SKILL.md for everything else".

**File:** [app.py:991-1055](../authoring/app.py#L991-L1055).

#### G5. `notes` silently dropped

Authors can write free-form `notes` in the describe UI. The index
helpers don't carry it through. Either include it (cheap, useful as
human-authored override hints — "do not use for technical content")
or document in the UI that notes are private scratch.

**Files:** same helpers as G1.

#### G6. No top-level summary in `index.json`

The agent gets a flat array. A 5-line preamble would help orientation
for larger workspaces:

```json
{
  "summary": {
    "templates": 24,
    "assets": 142,
    "decks": ["deckA", "deckB", "deckC"],
    "feels": {"formal": 14, "punchy": 6, "data-dense": 4}
  },
  "templates": [...],
  "assets": [...]
}
```

Cheap to compute, lets the agent calibrate strategy (broad vs
narrow KB) before scanning entries.

#### G7. Filter scope ≠ compose scope (worth documenting)

The bundle is filtered, but
[`_stage_compose_bundle`](../authoring/app.py#L1103-L1171) stages
the **full** KB when `compose run` executes. Picks the agent makes
from the filtered set always resolve; out-of-set ids the agent
hallucinates fail. Worth one line in `brief.md` so the agent
doesn't assume the bundle is the full universe.

---

## 2. In-bundle helper scripts

### Problem

If you imagine being the model: bundle arrives, code execution
available, no `reader.py`, no template binaries — just text + JSON +
theme files. To pick well across 50+ templates and 200+ assets, the
agent **will** write ad-hoc Python. Common ways that goes wrong:

- Treats `suitable_for` as string (it's a list).
- Treats `colors` as string (it's a list); confuses with `feel`
  (which IS a string).
- Filters on `interpretation` (SKILL.md says don't — it's info-only).
- Emits `{"text": "...", "color_role": "accent"}` thinking styling
  will render (SKILL.md says it's degraded today).
- Forgets the bullets-no-glyphs rule.
- Aspect ratios are `"16:9"` strings — float math without parsing
  fails.
- Theme color roles vs raw hex — agent reaches for
  `theme_colors.primary` not knowing aliases exist in theme.yaml.

A small set of prewritten scripts shipped *inside the bundle*
eliminates these by making "the right way" the easy way.

### Proposed helpers (CLI form so agents can shell out)

All read-only. All take `--json` flags for machine-parseable output.
Live under `helpers/` in the bundle. Each one self-contained
(stdlib + pyyaml only — pyyaml already in agent runtimes).

#### `helpers/kb_summary.py`

Prints faceted overview: counts per `feel`, `suitable_for`, `kind`,
per-deck breakdown, slot-budget distribution. Lets the agent pick a
strategy before scanning entries.

```bash
python helpers/kb_summary.py
python helpers/kb_summary.py --json
```

#### `helpers/kb_filter.py`

Catalog query with the right list/string semantics baked in. Returns
template *or* asset entries matching all criteria (AND across keys,
OR within a key via `|`).

```bash
python helpers/kb_filter.py templates --feel formal --suitable_for opener
python helpers/kb_filter.py assets --kind photo --feel warm --colors navy
python helpers/kb_filter.py templates --feel "warm|formal" --json
```

Mirrors `reader.py list` semantics so the agent doesn't have to
re-learn a second filter dialect.

#### `helpers/kb_inspect.py`

Denormalized view of a single template or asset. For a template:
inlines each inventory atom's full description so the agent can
score "does this template's anatomy match my brief" in one read.

```bash
python helpers/kb_inspect.py <id>
python helpers/kb_inspect.py deckA_01 --json
```

Closes G3 from the agent's side even if the index itself stays
normalized.

#### `helpers/kb_lint.py`

**The single most valuable helper.** Pre-flight validator for a
draft plan. Catches:

- Slot ids that don't exist on the referenced template.
- `max_chars` overflows (per slot, with overage count).
- Leading bullet glyphs in bullet slots or text-as-bullets slots.
- Asset ids referenced but not in `index.json`.
- Compose-mode entries that would currently be skipped (warn).
- Slot value shapes that are accepted-but-degraded (warn so the
  agent makes a conscious call).
- Required slots left unfilled.

```bash
python helpers/kb_lint.py < plan.json
python helpers/kb_lint.py plan.json --json
```

Exit code = number of errors. The agent can loop "draft → lint →
revise" without ever needing the user to round-trip through the
Compose UI.

#### `helpers/kb_themes.py`

Dumps every `decks/<deck>/theme.yaml` as a single table — palette +
aliases + fonts + aspect — so the agent can pick templates whose
theme is closest to brand policy.

```bash
python helpers/kb_themes.py
python helpers/kb_themes.py --resolve-role accent --for-deck deckA
```

The `--resolve-role` form returns the resolved hex for a role token,
saving the agent from re-implementing the alias chain.

#### `helpers/kb_budget.py`

One specific check separated out because it's the most common gotcha:
"will this string fit slot X on template Y?" Returns `(ok, used,
max)` for one slot, or a table for all slots in a plan.

```bash
python helpers/kb_budget.py <template_id> <slot_id> "draft text…"
```

(Could fold into `kb_lint` — keep separate iff the model finds itself
iterating on text for one slot at a time.)

### Bundle layout after these land

```
prompt-bundle-<ts>.zip
├── SKILL.md
├── brand.md
├── index.json
├── brief.md
├── decks/<deck>/theme.yaml
├── templates/<id>/preview.png   ← new, gated by checkbox (G2)
└── helpers/
    ├── README.md                short — what each script does, exit codes
    ├── kb_summary.py
    ├── kb_filter.py
    ├── kb_inspect.py
    ├── kb_lint.py
    ├── kb_themes.py
    └── kb_budget.py
```

### Design constraints for the helpers

- **No filesystem writes.** Bundle is read-only from the agent's POV;
  the only output is the plan, which the agent prints/saves itself.
- **Stdlib + pyyaml only.** No python-pptx — the bundle has no .pptx
  binaries to operate on. Rendering is the user's machine job.
- **`--json` everywhere.** Pretty-printed default for the agent's
  own reading, `--json` for parsing.
- **Mirror SKILL.md vocab exactly.** Same filter keys, same
  semantics for `|` OR and the `none` literal.
- **Stable exit codes.** 0 = clean / matches found, 1 = no matches
  / errors found, 2 = bad input. Lets the agent chain.
- **One file each, ≤200 lines.** Trivial for the agent to read +
  understand if it ever wants to verify behavior.

### What we explicitly should NOT ship as helpers

- A render helper — bundle has no template binaries.
- Anything that mutates files.
- A "search the user's workspace" helper — bundle is closed-world.
- A wrapper that calls an external API (would require keys).
- A "smart picker" that hides the catalog from the agent — defeats
  the point. Helpers expose data, agent decides.

---

## Suggested implementation order

1. **G4** (trim brief.md) — pure deletion, no behavior change, lets
   later changes land cleanly.
2. **G1** (add `deck` to index entries) — 2-line change in two
   helpers + test.
3. **kb_lint.py** — biggest single quality win for the agent.
4. **kb_filter.py + kb_inspect.py + kb_summary.py** — together
   cover the planning loop.
5. **G3** (inline inventory atoms) — nice-to-have once helpers
   exist; helpers make it less urgent.
6. **G6** (index summary block) — once we have decks per entry.
7. **G2** (previews, behind a checkbox) — biggest bundle-size
   impact, do last and measure.
8. **kb_themes.py + kb_budget.py** — secondary helpers.
9. **G5** (notes pass-through) — decide policy first (private vs
   public).
