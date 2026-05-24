#!/usr/bin/env python3
"""Faceted overview of the bundle catalog.

First thing to read for an unfamiliar bundle — calibrate scope before
scanning entries. Counts templates + assets per major facet.
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kb_common import default_bundle_root, load_index, load_themes


def _facet(items: list[dict], field: str) -> Counter:
    c: Counter = Counter()
    for it in items:
        v = it.get(field)
        if v is None or v == "" or v == []:
            c["(none)"] += 1
        elif isinstance(v, list):
            for x in v:
                if isinstance(x, str) and x:
                    c[x] += 1
        elif isinstance(v, str):
            c[v] += 1
    return c


def _print_facet(c: Counter, top: int | None = None) -> None:
    items = sorted(c.items(), key=lambda kv: (-kv[1], kv[0]))
    if top is not None:
        items = items[:top]
    if not items:
        print("  (empty)")
        return
    width = max(len(k) for k, _ in items)
    for k, v in items:
        print(f"  {k:<{width}}  {v}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--bundle", type=Path, default=None,
                        help="bundle root (default: helpers/..)")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    bundle = args.bundle or default_bundle_root(__file__)
    index = load_index(bundle)
    themes = load_themes(bundle)
    templates = index.get("templates", [])
    assets = index.get("assets", [])

    facets = {
        "templates": {
            "feel": _facet(templates, "feel"),
            "suitable_for": _facet(templates, "suitable_for"),
        },
        "assets": {
            "kind": _facet(assets, "kind"),
            "feel": _facet(assets, "feel"),
            "composition": _facet(assets, "composition"),
            "suitable_for": _facet(assets, "suitable_for"),
            "colors": _facet(assets, "colors"),
            "scope": _facet(assets, "scope"),
        },
    }

    if args.json:
        payload = {
            "counts": {
                "templates": len(templates),
                "assets": len(assets),
                "decks_with_theme": len(themes),
            },
            "themes": list(themes.keys()),
            "templates": {k: dict(v) for k, v in facets["templates"].items()},
            "assets": {k: dict(v) for k, v in facets["assets"].items()},
        }
        print(json.dumps(payload, indent=2, ensure_ascii=False))
        return 0

    print(f"Templates: {len(templates)}  •  "
          f"Assets: {len(assets)}  •  "
          f"Decks with theme: {len(themes)}")
    if themes:
        print(f"Decks: {', '.join(themes.keys())}")
    print()
    print("== Templates ==")
    print("feel:")
    _print_facet(facets["templates"]["feel"])
    print("suitable_for:")
    _print_facet(facets["templates"]["suitable_for"])
    print()
    print("== Assets ==")
    for f in ("kind", "feel", "composition", "suitable_for"):
        print(f"{f}:")
        _print_facet(facets["assets"][f])
    print("colors (top 20):")
    _print_facet(facets["assets"]["colors"], top=20)
    print("scope (top 20):")
    _print_facet(facets["assets"]["scope"], top=20)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
