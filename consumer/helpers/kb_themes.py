#!/usr/bin/env python3
"""Dump per-deck themes and resolve role tokens.

Each ingested deck ships `decks/<deck>/theme.yaml` with its palette,
alias map, fonts, and aspect. This script makes that data scannable
without reading every theme file individually.

Usage:
    python helpers/kb_themes.py
    python helpers/kb_themes.py --json
    python helpers/kb_themes.py --resolve-role accent --for-deck <deck>

Exit codes: 0 ok, 1 deck/theme not found, 2 bad input.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kb_common import default_bundle_root, load_themes, resolve_role


def _print_theme(deck: str, t: dict) -> None:
    palette = t.get("palette") or {}
    aliases = t.get("aliases") or {}
    fonts = t.get("fonts") or {}
    aspect = t.get("aspect") or "-"
    print(f"== {deck} ==  aspect: {aspect}")
    if fonts:
        print("  fonts:    " + ", ".join(f"{k}={v}" for k, v in fonts.items()))
    if aliases:
        print("  aliases:  "
              + ", ".join(f"{k}→{v}" for k, v in aliases.items()))
    if palette:
        print("  palette:")
        for k, v in palette.items():
            print(f"    {k:<10} {v}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--bundle", type=Path, default=None)
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--resolve-role", default=None,
                        help="role token to resolve "
                             "(primary|accent|text|background|<scheme>)")
    parser.add_argument("--for-deck", default=None,
                        help="when resolving, which deck's theme to use")
    args = parser.parse_args()

    bundle = args.bundle or default_bundle_root(__file__)
    themes = load_themes(bundle)

    if args.resolve_role is not None:
        if not args.for_deck:
            print("--resolve-role requires --for-deck", file=sys.stderr)
            return 2
        theme = themes.get(args.for_deck)
        if theme is None:
            print(f"no theme for deck {args.for_deck!r} "
                  f"(known: {sorted(themes.keys())})", file=sys.stderr)
            return 1
        hex_ = resolve_role(theme, args.resolve_role)
        if hex_ is None:
            print(f"role {args.resolve_role!r} did not resolve "
                  f"against deck {args.for_deck!r}", file=sys.stderr)
            return 1
        if args.json:
            print(json.dumps({
                "deck": args.for_deck,
                "role": args.resolve_role,
                "hex": hex_,
            }))
        else:
            print(hex_)
        return 0

    if not themes:
        print("(no themes in this bundle)")
        return 1

    if args.json:
        print(json.dumps(themes, indent=2, ensure_ascii=False))
        return 0

    for deck, t in themes.items():
        _print_theme(deck, t)
        print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
