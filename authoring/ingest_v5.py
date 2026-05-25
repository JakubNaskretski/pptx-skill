"""v5 redesign — structural-skeleton digest (self-contained).

This module is intentionally a single file so v5 is removable as one
delete + a tiny hook-removal in cli.py if the redesign doesn't pan
out. Until phase F flips the build flag, v5 outputs are purely
additive: v4 slide.yaml / theme.yaml continue to be written under
workspace/decks/<deck>/ untouched, and v5 writes alongside under
workspace/themes/<deck>/ and workspace/skeletons/<deck>_<NN>/.

See REDESIGN.md (root) for the architecture; phase callouts in
function docstrings reference sub-phases B1-B5 + C1.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import yaml


# ---------------------------------------------------------------------------
# Local helpers — duplicated from cli.py to keep this module
# import-cycle-free and easy to delete as a unit.
# ---------------------------------------------------------------------------


def _write_yaml(path: Path, data: dict) -> None:
    path.write_text(
        yaml.safe_dump(data, sort_keys=False, allow_unicode=True, width=100),
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# B1 — Theme extraction
# ---------------------------------------------------------------------------


def digest_theme(
    original_path: Path,
    deck_stem: str,
    theme_v4: dict,
    themes_root: Path,
) -> dict:
    """Write workspace/themes/<deck>/theme.yaml + master.pptx.

    Reads from the v4 theme dict (which already carries palette + fonts
    + aliases extracted by cli.extract_deck_theme) and emits the v5
    schema: semantic palette roles (primary / accent / text_default /
    background) resolved via the v4 aliases, fonts, master_pptx and
    preview references, and an empty decorations array (B3 populates).

    master.pptx is currently a copy of original.pptx. Proper master-only
    extraction (drop slides, keep masters + layouts + theme + media) is
    a future refinement — the phase-E build engine can strip slides at
    build time, so this is correctness-equivalent, just larger than the
    minimal artifact.

    Returns the written v5 theme dict (for callers that want to log).
    """
    theme_dir = themes_root / deck_stem
    theme_dir.mkdir(parents=True, exist_ok=True)

    master_pptx = theme_dir / "master.pptx"
    _extract_master_pptx(original_path, master_pptx)

    palette_v4 = theme_v4.get("palette") or {}
    aliases = theme_v4.get("aliases") or {}

    def _resolve(role: str) -> str:
        slot = aliases.get(role, "")
        return palette_v4.get(slot, "") if slot else ""

    palette_v5 = {
        "primary": _resolve("primary"),
        "accent": _resolve("accent"),
        "text_default": _resolve("text"),
        "background": _resolve("background"),
    }
    # Drop empties so partial extractions are visibly partial in the
    # YAML rather than misleadingly defaulting to "".
    palette_v5 = {k: v for k, v in palette_v5.items() if v}

    fonts = {k: v for k, v in (theme_v4.get("fonts") or {}).items() if v}

    out = {
        "id": deck_stem,
        "palette": palette_v5,
        "fonts": fonts,
        "master_pptx": "master.pptx",
        "preview": "preview.png",   # rendered separately; file may be absent
        "decorations": [],          # populated in B3
    }
    _write_yaml(theme_dir / "theme.yaml", out)
    return out


def _extract_master_pptx(src: Path, dst: Path) -> None:
    """Copy the original deck as master.pptx.

    Correctness-equivalent placeholder for the eventual master-only
    extraction (strip all slides, keep masters/layouts/theme/media).
    Build engine in phase E can handle stripping at build time, so a
    full copy works today — it just ships more bytes than necessary.
    Track via REDESIGN.md and tighten when phase E lands.
    """
    shutil.copyfile(src, dst)
