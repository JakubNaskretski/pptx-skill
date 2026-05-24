#!/usr/bin/env python3
"""Pre-flight validator for a draft plan.json.

Catches the most common errors before the user round-trips through the
Compose UI: bad template ids, bad slot ids, max_chars overflows, bullet-
glyph leakage, asset ids not in the index, plus warnings for
accepted-but-degraded shapes and compose-mode entries (which the current
engine skips with a warning).

Usage:
    python helpers/kb_lint.py < plan.json
    python helpers/kb_lint.py plan.json
    python helpers/kb_lint.py plan.json --json

Exit codes: 0 if clean (no errors; warnings still allowed),
1 if errors, 2 on bad input.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kb_common import (
    asset_ref_id,
    assets_by_id,
    default_bundle_root,
    extract_text,
    has_leading_bullet_glyph,
    is_degraded_shape,
    load_index,
    templates_by_id,
)


def _slot_lookup(tpl: dict) -> dict[str, dict]:
    return {s["id"]: s for s in (tpl.get("slots") or []) if s.get("id")}


def _lint_template_entry(entry: dict, idx: int, tpls: dict[str, dict],
                         assets: dict[str, dict]) -> list[dict]:
    out: list[dict] = []
    tid = entry.get("template")
    if not isinstance(tid, str):
        out.append({"level": "error", "entry": idx,
                    "msg": "missing or non-string 'template' field"})
        return out
    tpl = tpls.get(tid)
    if tpl is None:
        out.append({"level": "error", "entry": idx,
                    "msg": f"template id {tid!r} not in index"})
        return out
    slots_def = _slot_lookup(tpl)
    plan_slots = entry.get("slots") or {}
    if not isinstance(plan_slots, dict):
        out.append({"level": "error", "entry": idx, "template": tid,
                    "msg": "'slots' must be an object"})
        return out
    for sid, sval in plan_slots.items():
        sdef = slots_def.get(sid)
        if sdef is None:
            out.append({"level": "error", "entry": idx, "template": tid,
                        "slot": sid,
                        "msg": f"slot id {sid!r} not declared on template "
                               f"(declared: {sorted(slots_def.keys())})"})
            continue
        out.extend(_lint_slot_value(idx, tid, sid, sdef, sval, assets))
    return out


def _lint_slot_value(idx: int, tid: str, sid: str, sdef: dict, sval,
                     assets: dict[str, dict]) -> list[dict]:
    out: list[dict] = []
    kind = sdef.get("kind", "text")
    max_chars = sdef.get("max_chars")
    max_items = sdef.get("max_items")
    where = {"entry": idx, "template": tid, "slot": sid}

    if is_degraded_shape(sval):
        out.append({**where, "level": "warning",
                    "msg": "shape uses color_role/font_role/bold/runs/"
                           "recolor — accepted but degraded to plain in the "
                           "current build"})

    if kind == "image":
        if not (isinstance(sval, str) or isinstance(sval, dict)):
            out.append({**where, "level": "error",
                        "msg": "image slot value must be an asset id string "
                               "or an {asset, ...} object"})
            return out
        aid = asset_ref_id(sval)
        if aid is None:
            out.append({**where, "level": "error",
                        "msg": "image slot value did not resolve to an "
                               "asset_<id> reference"})
        elif aid not in assets:
            out.append({**where, "level": "error",
                        "msg": f"asset id {aid!r} not in index"})
        return out

    if kind == "bullets":
        if not isinstance(sval, list):
            out.append({**where, "level": "error",
                        "msg": "bullets slot value must be an array of strings"})
            return out
        if max_items is not None and len(sval) > max_items:
            out.append({**where, "level": "error",
                        "msg": f"bullets count {len(sval)} exceeds max_items "
                               f"{max_items}"})
        for i, item in enumerate(sval):
            if not isinstance(item, str):
                out.append({**where, "level": "error",
                            "msg": f"bullets[{i}] is not a string"})
                continue
            if has_leading_bullet_glyph(item):
                out.append({**where, "level": "error",
                            "msg": f"bullets[{i}] starts with a bullet glyph "
                                   f"— template applies bullets via layout; "
                                   f"do not prepend them"})
        return out

    if kind == "table":
        # accept list-of-lists or {"cells": [[...]]}
        cells = (sval.get("cells") if isinstance(sval, dict) else sval)
        if not (isinstance(cells, list)
                and all(isinstance(r, list) for r in cells)):
            out.append({**where, "level": "error",
                        "msg": "table slot value must be list-of-lists "
                               "(or {cells: [[...]]})"})
        return out

    # default: text-shaped slot
    text = extract_text(sval)
    if text is None:
        out.append({**where, "level": "error",
                    "msg": f"text slot value has unrecognized shape: "
                           f"{type(sval).__name__}"})
        return out
    if has_leading_bullet_glyph(text):
        out.append({**where, "level": "error",
                    "msg": "text starts with a bullet glyph — template "
                           "handles bullets via layout; do not prepend"})
    if max_chars is not None and len(text) > max_chars:
        out.append({**where, "level": "error",
                    "msg": f"text length {len(text)} exceeds max_chars "
                           f"{max_chars} (overage: "
                           f"{len(text) - max_chars})"})
    return out


def _lint_compose_entry(entry: dict, idx: int,
                        assets: dict[str, dict]) -> list[dict]:
    out: list[dict] = []
    out.append({"entry": idx, "level": "warning",
                "msg": "compose-mode entry — the current engine SKIPS these "
                       "with a warning (full support in Phase D). "
                       "Still validating shape for forward compatibility."})
    shapes = entry.get("shapes")
    if not isinstance(shapes, list):
        out.append({"entry": idx, "level": "error",
                    "msg": "compose entry missing 'shapes' array"})
        return out
    for i, sh in enumerate(shapes):
        if not isinstance(sh, dict):
            out.append({"entry": idx, "shape": i, "level": "error",
                        "msg": "shape must be an object"})
            continue
        kind = sh.get("kind")
        if kind == "text":
            if not isinstance(sh.get("value"), str):
                out.append({"entry": idx, "shape": i, "level": "error",
                            "msg": "text shape needs string 'value'"})
        else:
            atom = sh.get("atom")
            if not isinstance(atom, str) or not atom.startswith("asset_"):
                out.append({"entry": idx, "shape": i, "level": "error",
                            "msg": "non-text shape needs 'atom: asset_<id>'"})
            elif atom not in assets:
                out.append({"entry": idx, "shape": i, "level": "error",
                            "msg": f"atom {atom!r} not in index"})
        for axis in ("x", "y", "w", "h"):
            v = sh.get(axis)
            if v is None:
                continue
            if not isinstance(v, (int, float)) or not (0 <= v <= 1):
                out.append({"entry": idx, "shape": i, "level": "warning",
                            "msg": f"shape {axis}={v} outside [0,1] — "
                                   f"geometry is fractional"})
    return out


def lint(plan, index) -> list[dict]:
    results: list[dict] = []
    if not isinstance(plan, list):
        return [{"level": "error",
                 "msg": "plan must be a JSON array at the top level"}]
    tpls = templates_by_id(index)
    assets = assets_by_id(index)
    for idx, entry in enumerate(plan):
        if not isinstance(entry, dict):
            results.append({"entry": idx, "level": "error",
                            "msg": "entry must be an object"})
            continue
        if entry.get("compose") is True:
            results.extend(_lint_compose_entry(entry, idx, assets))
        elif "template" in entry:
            results.extend(_lint_template_entry(entry, idx, tpls, assets))
        else:
            results.append({"entry": idx, "level": "error",
                            "msg": "entry must have either 'template' or "
                                   "'compose: true'"})
    return results


def _read_plan(arg_path: str | None) -> object:
    if arg_path:
        return json.loads(Path(arg_path).read_text(encoding="utf-8"))
    return json.loads(sys.stdin.read())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("plan", nargs="?", help="path to plan.json (or stdin)")
    parser.add_argument("--bundle", type=Path, default=None)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    try:
        plan = _read_plan(args.plan)
    except (OSError, json.JSONDecodeError) as e:
        print(f"bad input: {e}", file=sys.stderr)
        return 2

    bundle = args.bundle or default_bundle_root(__file__)
    index = load_index(bundle)
    results = lint(plan, index)

    n_err = sum(1 for r in results if r.get("level") == "error")
    n_warn = sum(1 for r in results if r.get("level") == "warning")

    if args.json:
        print(json.dumps({
            "errors": n_err, "warnings": n_warn, "results": results,
        }, indent=2, ensure_ascii=False))
    else:
        if not results:
            print("clean — no errors or warnings")
        else:
            for r in results:
                loc_bits = []
                for k in ("entry", "template", "slot", "shape"):
                    if k in r:
                        loc_bits.append(f"{k}={r[k]}")
                loc = " ".join(loc_bits)
                print(f"[{r.get('level','?')}] {loc}: {r.get('msg','')}")
            print(f"\n{n_err} error(s), {n_warn} warning(s)")
    return 1 if n_err else 0


if __name__ == "__main__":
    raise SystemExit(main())
