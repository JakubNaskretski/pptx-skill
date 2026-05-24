"""Shared loaders + filter primitives for the kb_* helper scripts.

Lives alongside the helpers in the prompt bundle. Stdlib + pyyaml.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

try:
    import yaml  # only used by theme loaders
except ImportError:  # pragma: no cover
    yaml = None


# --- bundle loaders --------------------------------------------------------

def default_bundle_root(script_file: str) -> Path:
    """A helper script's parent's parent — i.e. the bundle root."""
    return Path(script_file).resolve().parent.parent


def load_index(bundle_root: Path) -> dict:
    p = bundle_root / "index.json"
    if not p.exists():
        raise SystemExit(f"index.json not found at {p}")
    return json.loads(p.read_text(encoding="utf-8"))


def load_themes(bundle_root: Path) -> dict:
    """Return {deck_name: theme_dict}. Empty if pyyaml missing or no themes."""
    if yaml is None:
        return {}
    decks_dir = bundle_root / "decks"
    if not decks_dir.exists():
        return {}
    out = {}
    for theme_yaml in sorted(decks_dir.glob("*/theme.yaml")):
        try:
            out[theme_yaml.parent.name] = (
                yaml.safe_load(theme_yaml.read_text(encoding="utf-8")) or {}
            )
        except Exception:
            continue
    return out


# --- filter primitives -----------------------------------------------------

def parse_filter_value(s: str) -> list[str]:
    """'a|b|c' -> ['a','b','c']. 'none' is a sentinel meaning empty/missing."""
    return [v.strip() for v in s.split("|") if v.strip()]


def entry_matches_filter(entry: dict, field: str, allowed: list[str]) -> bool:
    """One AND-clause: does `entry[field]` match any value in `allowed`?

    - 'none' in allowed matches missing/empty (None, "", [])
    - List-valued fields match if ANY element is in allowed
    - String-valued fields match by equality
    """
    v = entry.get(field)
    none_ok = "none" in allowed
    if v is None or v == "" or v == []:
        return none_ok
    if isinstance(v, list):
        return any(x in allowed for x in v)
    if isinstance(v, str):
        return v in allowed
    return False


TEXT_SEARCH_FIELDS_TEMPLATE = ("intent", "interpretation", "layout")
TEXT_SEARCH_FIELDS_ASSET = ("subject", "depicts", "interpretation")


def text_search(entry: dict, query: str, fields: tuple[str, ...]) -> bool:
    q = query.lower()
    for f in fields:
        v = entry.get(f)
        if isinstance(v, str) and q in v.lower():
            return True
        if isinstance(v, list):
            for x in v:
                if isinstance(x, str) and q in x.lower():
                    return True
    return False


# --- entry lookup ----------------------------------------------------------

def find_entry(index: dict, entry_id: str) -> tuple[str, dict | None]:
    """Return ('template'|'asset', dict) or ('', None) if not found."""
    for t in index.get("templates", []):
        if t.get("id") == entry_id:
            return "template", t
    for a in index.get("assets", []):
        if a.get("id") == entry_id:
            return "asset", a
    return "", None


def assets_by_id(index: dict) -> dict[str, dict]:
    return {a["id"]: a for a in index.get("assets", []) if a.get("id")}


def templates_by_id(index: dict) -> dict[str, dict]:
    return {t["id"]: t for t in index.get("templates", []) if t.get("id")}


def load_user_assets(bundle_root: Path) -> dict[str, dict]:
    """Read user_assets/manifest.json if present; return {id: entry}.

    These are user-supplied attachments to the request — same id format
    as catalog assets but listed separately (no descriptions). Empty
    dict if the bundle has none.
    """
    p = bundle_root / "user_assets" / "manifest.json"
    if not p.exists():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return data.get("assets") or {}


# --- plan-side helpers (used by kb_lint, kb_budget) ------------------------

LEADING_BULLET_RE = re.compile(r"^\s*[•\-*–—]\s+")


def has_leading_bullet_glyph(s: Any) -> bool:
    return isinstance(s, str) and bool(LEADING_BULLET_RE.match(s))


def extract_text(value: Any) -> str | None:
    """Pull text content from a polymorphic text-slot value.

    Returns None for shapes we can't read as text (e.g. asset refs).
    """
    if isinstance(value, str):
        return value
    if isinstance(value, dict):
        if isinstance(value.get("text"), str):
            return value["text"]
        runs = value.get("runs")
        if isinstance(runs, list):
            return "".join(
                r.get("text", "") for r in runs if isinstance(r, dict)
            )
    return None


def is_asset_ref(value: Any) -> bool:
    if isinstance(value, str) and value.startswith("asset_"):
        return True
    if isinstance(value, dict):
        a = value.get("asset")
        if isinstance(a, str) and a.startswith("asset_"):
            return True
    return False


def asset_ref_id(value: Any) -> str | None:
    if isinstance(value, str) and value.startswith("asset_"):
        return value
    if isinstance(value, dict):
        a = value.get("asset")
        if isinstance(a, str) and a.startswith("asset_"):
            return a
    return None


def is_degraded_shape(value: Any) -> bool:
    """Detect slot value shapes that SKILL.md says are accepted-but-degraded
    in the current build (styling/runs/recolor). Useful for lint warnings.
    """
    if not isinstance(value, dict):
        return False
    keys = set(value.keys())
    return bool(keys & {"color_role", "font_role", "bold", "runs", "recolor"})


# --- theme / role resolution -----------------------------------------------

ROLE_FALLBACKS = {
    "primary": "dk1",
    "accent": "accent1",
    "text": "dk1",
    "background": "lt1",
}


def resolve_role(theme: dict, role: str) -> str | None:
    """role token -> hex from a theme.yaml dict.

    Looks at aliases first; falls back to OOXML scheme slot names; else
    treats `role` as a literal palette key.
    """
    palette = theme.get("palette") or {}
    aliases = theme.get("aliases") or {}
    target = aliases.get(role) or ROLE_FALLBACKS.get(role) or role
    return palette.get(target)
