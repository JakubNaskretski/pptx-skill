#!/usr/bin/env python3
"""Check a single slot's text budget without writing a full plan.

For iterating on one slot's copy ("does this title fit?"). `kb_lint.py`
covers the same check across a full plan; this is the one-shot variant.

Usage:
    python helpers/kb_budget.py <template_id> <slot_id> "draft text"
    python helpers/kb_budget.py <template_id> <slot_id> --stdin

Exit codes: 0 fits, 1 overflows, 2 bad input.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kb_common import default_bundle_root, load_index, templates_by_id


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("template_id")
    parser.add_argument("slot_id")
    parser.add_argument("text", nargs="?", default=None,
                        help="draft text (or use --stdin)")
    parser.add_argument("--stdin", action="store_true",
                        help="read draft text from stdin")
    parser.add_argument("--bundle", type=Path, default=None)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    if args.stdin and args.text is None:
        text = sys.stdin.read()
    elif args.text is not None:
        text = args.text
    else:
        print("provide text as a positional arg or --stdin", file=sys.stderr)
        return 2

    bundle = args.bundle or default_bundle_root(__file__)
    index = load_index(bundle)
    tpl = templates_by_id(index).get(args.template_id)
    if tpl is None:
        print(f"template {args.template_id!r} not in index", file=sys.stderr)
        return 2
    slot = next((s for s in (tpl.get("slots") or [])
                 if s.get("id") == args.slot_id), None)
    if slot is None:
        print(f"slot {args.slot_id!r} not declared on template "
              f"{args.template_id!r}", file=sys.stderr)
        return 2

    kind = slot.get("kind", "text")
    max_chars = slot.get("max_chars")
    used = len(text)

    if kind not in ("text", "bullets"):
        # bullets max_items would need a structured count; budget here is for
        # text/bullets-as-text. For bullets per-line, use kb_lint.
        if args.json:
            print(json.dumps({
                "ok": True, "kind": kind,
                "msg": "slot is not text-shaped; budget check skipped",
            }))
        else:
            print(f"slot kind={kind} — budget check skipped (use kb_lint "
                  f"for full validation)")
        return 0

    if max_chars is None:
        if args.json:
            print(json.dumps({
                "ok": True, "used": used, "max_chars": None,
                "msg": "slot has no max_chars declared",
            }))
        else:
            print(f"used={used}  max_chars=(none declared)  →  ok")
        return 0

    ok = used <= max_chars
    overage = max(0, used - max_chars)
    if args.json:
        print(json.dumps({
            "ok": ok, "used": used, "max_chars": max_chars,
            "overage": overage,
        }))
    else:
        verdict = "ok" if ok else f"OVER by {overage}"
        print(f"used={used}  max_chars={max_chars}  →  {verdict}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
