#!/usr/bin/env python3
"""Denormalized view of a single template or asset.

For a template, this inlines each `inventory[]` atom's full description
so you can score "does this template's anatomy match my brief" in one
read instead of cross-referencing the assets array per atom.

Exit codes: 0 if found, 1 if no entry by that id, 2 on bad input.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kb_common import (
    assets_by_id,
    default_bundle_root,
    find_entry,
    load_index,
)


ATOM_INLINE_FIELDS = ("kind", "subject", "feel", "composition", "colors")


def _denormalize_template(t: dict, assets_lookup: dict[str, dict]) -> dict:
    out = dict(t)
    inv = []
    for atom_entry in t.get("inventory") or []:
        merged = dict(atom_entry)
        atom_id = atom_entry.get("atom")
        asset = assets_lookup.get(atom_id) if atom_id else None
        if asset:
            merged["atom_meta"] = {
                k: asset.get(k) for k in ATOM_INLINE_FIELDS if asset.get(k)
            }
        else:
            merged["atom_meta"] = None  # asset not in bundle (filtered out)
        inv.append(merged)
    if inv:
        out["inventory"] = inv
    return out


def _format_slot(s: dict) -> str:
    bits = [f"id={s.get('id')!r}", f"kind={s.get('kind','?')}"]
    for k in ("max_chars", "max_items", "aspect"):
        if k in s:
            bits.append(f"{k}={s[k]}")
    style = s.get("style") or {}
    if style:
        bits.append(f"style={json.dumps(style, ensure_ascii=False)}")
    return "  " + ", ".join(bits)


def _format_template(t: dict) -> str:
    lines = [
        f"template  {t.get('id')}",
        f"  intent:        {t.get('intent') or ''}",
        f"  feel:          {t.get('feel') or '-'}",
        f"  suitable_for:  {', '.join(t.get('suitable_for') or []) or '-'}",
        f"  layout:        {t.get('layout') or '-'}",
    ]
    tc = t.get("theme_colors") or {}
    if tc:
        lines.append("  theme_colors:  "
                     + ", ".join(f"{k}={v}" for k, v in tc.items()))
    fonts = t.get("fonts") or {}
    if fonts:
        lines.append("  fonts:         "
                     + ", ".join(f"{k}={v}" for k, v in fonts.items()))
    interp = t.get("interpretation")
    if interp:
        lines.append(f"  interpretation: {interp}")
    lines.append("  slots:")
    for s in t.get("slots") or []:
        lines.append(_format_slot(s))
    inv = t.get("inventory") or []
    if inv:
        lines.append("  inventory:")
        for i in inv:
            head = (f"    atom={i.get('atom')}  kind={i.get('kind','?')}  "
                    f"region={i.get('region','?')}  "
                    f"x={i.get('x')} y={i.get('y')} "
                    f"w={i.get('w')} h={i.get('h')}")
            lines.append(head)
            meta = i.get("atom_meta")
            if meta:
                inner = ", ".join(
                    f"{k}={v if not isinstance(v, list) else ','.join(map(str, v))}"
                    for k, v in meta.items()
                )
                lines.append(f"        ↳ {inner}")
            elif meta is None and i.get("atom"):
                lines.append("        ↳ (atom not in this bundle)")
    return "\n".join(lines)


def _format_asset(a: dict) -> str:
    lines = [
        f"asset  {a.get('id')}",
        f"  kind:          {a.get('kind') or '-'}",
        f"  subject:       {a.get('subject') or ''}",
    ]
    if a.get("depicts"):
        lines.append(f"  depicts:       {a['depicts']}")
    lines.append(f"  feel:          {a.get('feel') or '-'}")
    if a.get("composition"):
        lines.append(f"  composition:   {a['composition']}")
    lines.append("  colors:        "
                 + (", ".join(a.get("colors") or []) or "-"))
    if a.get("colors_hex"):
        lines.append("  colors_hex:    " + ", ".join(a["colors_hex"]))
    if a.get("scope"):
        lines.append("  scope:         " + ", ".join(a["scope"]))
    lines.append("  suitable_for:  "
                 + (", ".join(a.get("suitable_for") or []) or "-"))
    for blk in ("table", "chart", "shape", "smartart"):
        v = a.get(blk)
        if v:
            lines.append(f"  {blk}: {json.dumps(v, ensure_ascii=False)}")
    rt = a.get("recolor_targets")
    if rt:
        lines.append("  recolor_targets: " + ", ".join(rt))
    interp = a.get("interpretation")
    if interp:
        lines.append(f"  interpretation: {interp}")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("id", help="template id or asset id")
    parser.add_argument("--bundle", type=Path, default=None)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    bundle = args.bundle or default_bundle_root(__file__)
    index = load_index(bundle)
    kind, entry = find_entry(index, args.id)
    if entry is None:
        print(f"no entry with id {args.id!r}", file=sys.stderr)
        return 1
    if kind == "template":
        out = _denormalize_template(entry, assets_by_id(index))
        if args.json:
            print(json.dumps(out, indent=2, ensure_ascii=False))
        else:
            print(_format_template(out))
    else:
        if args.json:
            print(json.dumps(entry, indent=2, ensure_ascii=False))
        else:
            print(_format_asset(entry))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
