# Bundle helpers

Read-only Python utilities for working with this prompt bundle from an
agent that has code execution. Stdlib + pyyaml.

Run from the bundle root:

```bash
python helpers/kb_summary.py
python helpers/kb_filter.py templates --feel formal --suitable_for opener
python helpers/kb_filter.py assets --kind photo --text "team"
python helpers/kb_inspect.py <template_or_asset_id>
python helpers/kb_lint.py < plan.json
python helpers/kb_themes.py
python helpers/kb_themes.py --resolve-role accent --for-deck <deck>
python helpers/kb_budget.py <template_id> <slot_id> "draft text"
```

Each script accepts `--bundle <path>` to point at a bundle directory
other than the script's own parent (useful for testing). Most accept
`--json` for machine-parseable output.

## When to use which

| Step | Helper | Why |
|---|---|---|
| Calibrate scope on first read | `kb_summary` | Facet counts before scanning entries |
| Find candidate templates / assets | `kb_filter` | Same filter dialect as `reader.py list`; `--text` for free-text |
| Score one candidate | `kb_inspect` | Denormalized view; inlines `inventory[]` atoms with their descriptions |
| Sanity-check copy fit | `kb_budget` | One-shot max_chars check while iterating |
| Validate a draft plan | `kb_lint` | Catches slot id typos, max_chars overflows, bullet glyph leakage, missing assets, degraded-shape warnings, compose-mode skip warning |
| Brand-fit / color resolution | `kb_themes` | Per-deck palette + alias-aware role resolution |

## Exit codes

- `0` — clean (matches found, validation passed, etc.)
- `1` — empty result (no matches) or validation failed
- `2` — bad input (missing args, malformed JSON)

The agent can chain helpers based on exit codes.

## What these helpers will NOT do

- Write files or mutate the bundle.
- Render or compose .pptx — bundle has no template binaries; that's the
  user's `reader.py compose` step.
- Call external APIs or require network access.
- Hide the catalog behind a "smart picker" — helpers expose data, the
  agent decides.

## Filter dialect

Mirrors `reader.py list --filter`:

- Multiple `--field=value` flags AND together (across keys).
- Within one flag, `|` is OR: `--feel "warm|formal"`.
- The literal `none` matches items where the field is empty/missing:
  `--feel none` finds untagged items, `--feel "warm|none"` includes
  both warm and untagged.

`kb_filter --text "..."` is case-insensitive substring search across
`intent` / `interpretation` / `layout` (templates) or `subject` /
`depicts` / `interpretation` (assets). Don't structurally filter on
`interpretation` itself — it's info-only per SKILL.md.
