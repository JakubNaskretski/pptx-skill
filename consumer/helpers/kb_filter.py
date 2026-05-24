#!/usr/bin/env python3
"""Filter the catalog.

Mirrors `reader.py list --filter` semantics so the agent doesn't have to
learn a second filter dialect:
- `--feel "warm|formal"` — OR within a key
- `--suitable_for opener --feel formal` — AND across keys
- the literal value `none` matches items where the field is empty/missing
- `--text "thesis"` — case-insensitive substring across intent/subject/
  depicts/interpretation (asset-side) or intent/interpretation/layout
  (template-side)

Exit codes: 0 if matches found, 1 if none, 2 on bad input.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kb_common import (
    TEXT_SEARCH_FIELDS_ASSET,
    TEXT_SEARCH_FIELDS_TEMPLATE,
    default_bundle_root,
    entry_matches_filter,
    load_index,
    parse_filter_value,
    text_search,
)


TEMPLATE_FILTERS = ("feel", "suitable_for")
ASSET_FILTERS = ("kind", "feel", "composition", "suitable_for", "scope", "colors")


def _apply_filters(items: list[dict], raw_filters: dict[str, str],
                   text_query: str | None,
                   text_fields: tuple[str, ...]) -> list[dict]:
    out = []
    parsed = {k: parse_filter_value(v) for k, v in raw_filters.items() if v}
    for it in items:
        ok = True
        for field, allowed in parsed.items():
            if not entry_matches_filter(it, field, allowed):
                ok = False
                break
        if ok and text_query:
            ok = text_search(it, text_query, text_fields)
        if ok:
            out.append(it)
    return out


def _format_template(t: dict) -> str:
    slot_summary = ", ".join(
        f"{s.get('id')}:{s.get('kind','?')}" for s in (t.get("slots") or [])
    )
    return (
        f"{t.get('id')}  [{t.get('feel') or '-'}]  "
        f"suitable_for={','.join(t.get('suitable_for') or []) or '-'}\n"
        f"    intent: {t.get('intent') or ''}\n"
        f"    layout: {t.get('layout') or '-'}\n"
        f"    slots:  {slot_summary or '-'}"
    )


def _format_asset(a: dict) -> str:
    return (
        f"{a.get('id')}  [{a.get('kind') or '-'}]  "
        f"feel={a.get('feel') or '-'}  "
        f"suitable_for={','.join(a.get('suitable_for') or []) or '-'}\n"
        f"    subject: {a.get('subject') or ''}\n"
        f"    colors:  {','.join(a.get('colors') or []) or '-'}"
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = parser.add_subparsers(dest="kind", required=True)

    # Globals live on subparsers only — flags after the subcommand is the
    # natural order, and putting them on both parents would let the
    # subparser default (None / False) overwrite a value set on the top-
    # level parser.
    def _add_globals(p: argparse.ArgumentParser) -> None:
        p.add_argument("--bundle", type=Path, default=None)
        p.add_argument("--json", action="store_true")

    t = sub.add_parser("templates", help="filter templates")
    _add_globals(t)
    for f in TEMPLATE_FILTERS:
        t.add_argument(f"--{f}", default="", metavar="V|V|...")
    t.add_argument("--text", default="", metavar="QUERY")

    a = sub.add_parser("assets", help="filter assets")
    _add_globals(a)
    for f in ASSET_FILTERS:
        a.add_argument(f"--{f}", default="", metavar="V|V|...")
    a.add_argument("--text", default="", metavar="QUERY")

    args = parser.parse_args()
    bundle = args.bundle or default_bundle_root(__file__)
    index = load_index(bundle)

    if args.kind == "templates":
        raw = {f: getattr(args, f) for f in TEMPLATE_FILTERS}
        items = _apply_filters(
            index.get("templates", []), raw, args.text or None,
            TEXT_SEARCH_FIELDS_TEMPLATE,
        )
        formatter = _format_template
    else:
        raw = {f: getattr(args, f) for f in ASSET_FILTERS}
        items = _apply_filters(
            index.get("assets", []), raw, args.text or None,
            TEXT_SEARCH_FIELDS_ASSET,
        )
        formatter = _format_asset

    if args.json:
        print(json.dumps(items, indent=2, ensure_ascii=False))
    else:
        if not items:
            print(f"(no {args.kind} matched)")
        else:
            print(f"{len(items)} {args.kind} matched:\n")
            for it in items:
                print(formatter(it))
                print()

    return 0 if items else 1


if __name__ == "__main__":
    raise SystemExit(main())
