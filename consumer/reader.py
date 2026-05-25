"""pptx-skill consumer reader.

Three commands an agent can call to use a built skill bundle:

  python reader.py list [--filter key=value,key=value]
  python reader.py get <id>
  python reader.py compose <plan.json> <output.pptx>

All commands write JSON to stdout (compose also writes the deck).
No state. No vision required. Read SKILL.md for the contract.
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

import yaml
from pptx import Presentation


HERE = Path(__file__).resolve().parent


# ---------------------------------------------------------------------------
# Bundle layout helpers
# ---------------------------------------------------------------------------


def bundle_root() -> Path:
    return HERE


def load_index() -> dict:
    p = bundle_root() / "index.json"
    if not p.exists():
        raise SystemExit(f"index.json not found at {p}")
    return json.loads(p.read_text(encoding="utf-8"))


def template_dir(tid: str) -> Path:
    return bundle_root() / "templates" / tid


def asset_path(aid: str) -> Path | None:
    """Find the asset binary by id (asset_<sha8>) — any extension."""
    assets_dir = bundle_root() / "assets"
    if not assets_dir.exists():
        return None
    for cand in assets_dir.glob(f"{aid}.*"):
        if cand.suffix == ".yaml":
            continue
        return cand
    return None


def asset_meta_path(aid: str) -> Path:
    return bundle_root() / "assets" / f"{aid}.yaml"


# ---------------------------------------------------------------------------
# Filter
# ---------------------------------------------------------------------------


UNSET_SENTINEL = "none"


def parse_filter(s: str | None) -> dict[str, list[str]]:
    """Parse `k1=v1,k2=v2|v3` into {k1: [v1], k2: [v2, v3]}.

    Pipe (`|`) expresses OR within a key. The literal value `none`
    matches items where the field is missing or empty — useful to
    include un-tagged items (`feel=warm|none`) or to audit them
    explicitly (`feel=none`).
    """
    if not s:
        return {}
    out: dict[str, list[str]] = {}
    for chunk in s.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "=" not in chunk:
            raise SystemExit(f"bad --filter token (need key=value): {chunk!r}")
        k, v = chunk.split("=", 1)
        out[k.strip()] = [x.strip() for x in v.split("|") if x.strip()]
    return out


def _field_is_unset(got: Any) -> bool:
    if got is None:
        return True
    if isinstance(got, (str, list, dict)) and not got:
        return True
    return False


def matches_filter(item: dict, flt: dict[str, list[str]]) -> bool:
    for k, wants in flt.items():
        unset_ok = UNSET_SENTINEL in wants
        explicit = [w for w in wants if w != UNSET_SENTINEL]
        if k not in item or _field_is_unset(item.get(k)):
            if unset_ok:
                continue
            return False
        if not explicit:
            # Filter was `k=none` only and the field IS set → exclude.
            return False
        got = item[k]
        if isinstance(got, list):
            got_strs = [str(x) for x in got]
            if not any(w in got_strs for w in explicit):
                return False
        else:
            if str(got) not in explicit:
                return False
    return True


# ---------------------------------------------------------------------------
# Compose
# ---------------------------------------------------------------------------


def _find_shape_by_name(slide, name: str):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


_BULLET_PREFIX_RE = re.compile(r"^\s*[•·◦▪▫●◆■□*\-–—]\s+")


def _strip_bullet_prefix(text: str) -> str:
    """Strip leading bullet glyphs from each line.

    Templates render bullets via layout formatting; if the caller also
    prepended a glyph, the slide ends up with two bullets per line.
    """
    if not text:
        return text
    return "\n".join(_BULLET_PREFIX_RE.sub("", ln) for ln in text.split("\n"))


def _copy_run_color(src_font, dst_font) -> None:
    """Copy a run's font.color from src to dst, handling RGB + theme.

    ColorFormat assignment switches the colour type on assignment, so
    we try RGB first (most common) then theme-color. Silent fallback
    if neither resolves — leaves dst at its inherited default.
    """
    try:
        ctype = src_font.color.type
    except (AttributeError, ValueError):
        return
    if ctype is None:
        return
    try:
        rgb = src_font.color.rgb
        if rgb is not None:
            dst_font.color.rgb = rgb
            return
    except (AttributeError, ValueError):
        pass
    try:
        tc = src_font.color.theme_color
        if tc is not None:
            dst_font.color.theme_color = tc
    except (AttributeError, ValueError):
        pass


def _copy_run_font(src_font, dst_font) -> None:
    """Copy size/bold/italic/name/color from a template run to a new run."""
    try:
        if src_font.size is not None:
            dst_font.size = src_font.size
        if src_font.bold is not None:
            dst_font.bold = src_font.bold
        if src_font.italic is not None:
            dst_font.italic = src_font.italic
        if src_font.name is not None:
            dst_font.name = src_font.name
    except (AttributeError, ValueError):
        pass
    _copy_run_color(src_font, dst_font)


def _fill_text_shape(shape, value: str) -> None:
    """Replace the shape's text frame content with `value`, keeping its first
    paragraph's font/style (incl. colour) as the template."""
    if not shape.has_text_frame:
        return
    value = _strip_bullet_prefix(str(value))
    tf = shape.text_frame
    # Capture the first run's formatting cues if available.
    first_para = tf.paragraphs[0] if tf.paragraphs else None
    template_run = first_para.runs[0] if (first_para and first_para.runs) else None

    tf.clear()
    p = tf.paragraphs[0]
    if template_run is not None:
        # text_frame.clear() leaves an empty first paragraph — reuse it.
        run = p.add_run()
        run.text = value
        _copy_run_font(template_run.font, run.font)
    else:
        p.text = value


def _fill_bullets_shape(shape, values: list[str]) -> None:
    """Replace bullets in a text-frame placeholder. Each value becomes one
    paragraph (newlines within a value become soft breaks)."""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    first_para = tf.paragraphs[0] if tf.paragraphs else None
    template_run = first_para.runs[0] if (first_para and first_para.runs) else None
    template_level = first_para.level if first_para is not None else 0

    tf.clear()
    if not values:
        return
    first = True
    for value in values:
        lines = _strip_bullet_prefix(str(value)).split("\n")
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.level = template_level
        run = p.add_run()
        run.text = lines[0]
        if template_run is not None:
            _copy_run_font(template_run.font, run.font)
        # Sub-runs for additional lines within a bullet inherit the
        # same font + colour as the primary run. (FINDINGS A3.17 noted
        # that without this they fall back to paragraph defaults and
        # show visible drift mid-bullet. Sticking with "\n" in run.text
        # for the soft break — the proper <a:br/> rewrite is a Phase
        # D follow-up.)
        for extra in lines[1:]:
            sub = p.add_run()
            sub.text = "\n" + extra
            if template_run is not None:
                _copy_run_font(template_run.font, sub.font)


def _fill_table_shape(shape, cells: list[list]) -> list[str]:
    """Fill a table placeholder's cells from a list-of-lists.

    The agent supplies cells as `[[row0_col0, row0_col1, ...], ...]`.
    Cells beyond what the template table holds are dropped with a
    warning. Cells the agent doesn't provide are left untouched —
    that keeps decorative template rows (e.g. footer summary) intact
    when the agent only wants to overwrite a subset. Per-cell font /
    colour formatting is inherited from each cell's existing first
    run.

    Returns a list of non-fatal warning strings.
    """
    warnings: list[str] = []
    if not getattr(shape, "has_table", False):
        warnings.append(f"slot '{shape.name}': not a table shape; skipping table fill")
        return warnings
    table = shape.table
    rows = list(table.rows)
    nrows = len(rows)

    if len(cells) > nrows:
        warnings.append(
            f"slot '{shape.name}': got {len(cells)} table rows, "
            f"template has {nrows}; truncating extras"
        )

    for i, row_vals in enumerate(cells[:nrows]):
        row_cells = list(rows[i].cells)
        ncols = len(row_cells)
        if not isinstance(row_vals, list):
            warnings.append(
                f"slot '{shape.name}': row {i} is not a list (got {type(row_vals).__name__}); skipping"
            )
            continue
        if len(row_vals) > ncols:
            warnings.append(
                f"slot '{shape.name}': row {i} got {len(row_vals)} cols, "
                f"template has {ncols}; truncating extras"
            )
        for j, val in enumerate(row_vals[:ncols]):
            cell = row_cells[j]
            tf = cell.text_frame
            first_para = tf.paragraphs[0] if tf.paragraphs else None
            template_run = first_para.runs[0] if (first_para and first_para.runs) else None
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = _strip_bullet_prefix(str(val))
            if template_run is not None:
                _copy_run_font(template_run.font, run.font)

    return warnings


def _replace_image_shape_legacy(slide, shape, image_path: Path) -> None:
    """Remove the existing Picture shape and add a fresh one at the same
    geometry. Loses crop / border / shadow / rotation / transparency /
    effects / alt text. Used as a fallback when the in-place rewire
    can't find what it needs."""
    left = shape.left
    top = shape.top
    width = shape.width
    height = shape.height
    name = shape.name

    sp = shape._element
    sp.getparent().remove(sp)

    new_pic = slide.shapes.add_picture(
        str(image_path), left, top, width=width, height=height
    )
    new_pic.name = name


def _replace_image_shape(slide, shape, image_path: Path) -> None:
    """Swap a Picture shape's image while preserving everything else.

    Strategy: register the new image as a slide-part rel (use the
    public add_picture API as the registration mechanism, then
    immediately discard the temporary shape), then rewire the
    existing shape's <a:blip r:embed> at the new rel id. Crop,
    border, rotation, shadow, transparency, picture effects, alt
    text — all preserved.

    Falls back to remove+re-add (the v3 behaviour) if the existing
    shape isn't a standard blipFill picture or if anything else
    unexpected happens.
    """
    try:
        from pptx.oxml.ns import qn
    except ImportError:
        _replace_image_shape_legacy(slide, shape, image_path)
        return

    blipFill = shape._element.find(qn("p:blipFill"))
    if blipFill is None:
        _replace_image_shape_legacy(slide, shape, image_path)
        return
    blip = blipFill.find(qn("a:blip"))
    if blip is None:
        _replace_image_shape_legacy(slide, shape, image_path)
        return

    embed_attr = qn("r:embed")

    # Use add_picture as the registration vehicle: it imports the
    # image into the slide's rels properly across all python-pptx
    # versions. We grab the rel id, then remove the temporary shape
    # — the image part stays bound to the slide via the rel.
    try:
        tmp_pic = slide.shapes.add_picture(
            str(image_path),
            left=0, top=0,
            width=shape.width, height=shape.height,
        )
        tmp_blip = tmp_pic._element.find(".//" + qn("a:blip"))
        if tmp_blip is None:
            tmp_pic._element.getparent().remove(tmp_pic._element)
            _replace_image_shape_legacy(slide, shape, image_path)
            return
        new_rid = tmp_blip.get(embed_attr)
        tmp_pic._element.getparent().remove(tmp_pic._element)
    except Exception:
        _replace_image_shape_legacy(slide, shape, image_path)
        return

    if not new_rid:
        _replace_image_shape_legacy(slide, shape, image_path)
        return

    blip.set(embed_attr, new_rid)


def _degrade_styled_value(value: Any) -> tuple[Any, list[str]]:
    """Reduce a v4-styled slot value to its v3-compatible primitive.

    The agent can emit (per SKILL.md v4):
      - {"text": "...", "color_role": "...", "bold": true, ...}
      - {"runs": [{"text": "X", "bold": true}, {"text": " Y"}]}
      - {"asset": "asset_<id>", "recolor": {"#ff0000": "accent"}}

    Full handling (colour overrides, per-run formatting, image recolour)
    lands in Phase D — until then we extract the primitive payload and
    emit a one-line warning about what's being dropped. Existing
    string/list/asset_id values pass through unchanged.
    """
    if not isinstance(value, dict):
        return value, []

    warnings: list[str] = []
    if "runs" in value and isinstance(value["runs"], list):
        text = "".join(
            str(r.get("text", "")) for r in value["runs"] if isinstance(r, dict)
        )
        warnings.append(
            f"per-run formatting not yet honored — flattened to plain text {text!r}"
        )
        return text, warnings
    if "asset" in value:
        aid = str(value["asset"])
        ignored = sorted(k for k in value if k != "asset")
        if ignored:
            warnings.append(
                f"image overrides not yet honored — ignoring {ignored}"
            )
        return aid, warnings
    if "text" in value:
        text = str(value["text"])
        ignored = sorted(k for k in value if k != "text")
        if ignored:
            warnings.append(
                f"text styling not yet honored — ignoring {ignored}"
            )
        return text, warnings

    warnings.append(f"unrecognised slot value shape: {value!r}")
    return value, warnings


def _apply_slot_value(slide, slot_id: str, value: Any, kind_hint: str | None) -> list[str]:
    """Apply one slot. Returns a list of warning strings (not fatal)."""
    warnings: list[str] = []
    shape = _find_shape_by_name(slide, slot_id)
    if shape is None:
        warnings.append(f"slot '{slot_id}' not found on slide (no shape with that name)")
        return warnings

    # Table slots use a list-of-lists shape that the degrade-and-flatten
    # pre-pass below would corrupt — handle them up-front. We trust the
    # slot's kind hint, but also detect has_table on the shape as a
    # fallback when the meta is silent (pre-D3 templates).
    is_table_slot = kind_hint == "table" or (
        kind_hint is None and getattr(shape, "has_table", False)
    )
    if is_table_slot:
        cells = value
        if isinstance(value, dict):
            ignored = sorted(k for k in value if k != "cells")
            if ignored:
                warnings.append(
                    f"slot '{slot_id}': table slot only honors 'cells'; ignoring {ignored}"
                )
            cells = value.get("cells")
        if not isinstance(cells, list) or not all(isinstance(r, list) for r in cells):
            warnings.append(
                f"slot '{slot_id}': table slot expects list-of-lists "
                f"(got {type(value).__name__}); leaving template cells unchanged"
            )
            return warnings
        warnings.extend(_fill_table_shape(shape, cells))
        return warnings

    # v4: degrade styled/per-run/image-override dicts to v3 primitives.
    # Bullets lists may carry styled dicts per item — degrade each one.
    if isinstance(value, list):
        flat: list = []
        for item in value:
            primitive, item_warns = _degrade_styled_value(item)
            flat.append(primitive)
            warnings.extend(item_warns)
        value = flat
    else:
        value, scalar_warns = _degrade_styled_value(value)
        warnings.extend(scalar_warns)

    kind = kind_hint
    if kind is None:
        # Infer from shape / value.
        if isinstance(value, list):
            kind = "bullets"
        elif (
            getattr(shape, "shape_type", None) is not None
            and str(shape.shape_type).endswith("PICTURE")
        ) or (
            isinstance(value, str) and value.startswith("asset_")
        ):
            kind = "image"
        else:
            kind = "text"

    if kind == "image":
        aid = value if isinstance(value, str) else ""
        if not aid.startswith("asset_"):
            warnings.append(f"slot '{slot_id}': image slot expects asset_<id>, got {value!r}")
            return warnings
        bin_path = asset_path(aid)
        if bin_path is None:
            warnings.append(f"slot '{slot_id}': asset {aid} not found in bundle")
            return warnings
        _replace_image_shape(slide, shape, bin_path)
        return warnings

    if kind == "bullets":
        items = value if isinstance(value, list) else [str(value)]
        _fill_bullets_shape(shape, [str(v) for v in items])
        return warnings

    # text
    _fill_text_shape(shape, str(value))
    return warnings


# --- cross-deck slide copy --------------------------------------------------


# ---------------------------------------------------------------------------
# Deck-theme normalisation (v4 — D5)
# ---------------------------------------------------------------------------
#
# When we copy a foreign deck's slide / atom into the host package, any
# `<a:schemeClr val="..."/>` reference resolves against the HOST's theme
# at render time — so a "brand pink accent1" in the source might paint as
# a "navy accent1" in the host. We mitigate by remapping the small set of
# semantically-named slots (primary / accent / text / background) per the
# decks' aliases, so a shape that used "the source deck's accent" still
# uses "the host deck's accent" after the copy. Other clrScheme slots
# (accent2-6, dk3, …) we accept as-is — there is no canonical semantic
# mapping for them.

# Both names are valid in <a:schemeClr val>: theme-side (dk1/lt1/dk2/lt2)
# and use-site (tx1/bg1/tx2/bg2). Normalize both forms to the theme-side.
_SCHEME_NORMALIZE = {
    "tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2",
}


def _norm_scheme_slot(val: str) -> str:
    v = (val or "").lower()
    return _SCHEME_NORMALIZE.get(v, v)


def _load_deck_theme(deck_name: str) -> dict | None:
    """Load decks/<deck>/theme.yaml from the bundle; None if missing."""
    if not deck_name:
        return None
    p = bundle_root() / "decks" / deck_name / "theme.yaml"
    if not p.exists():
        return None
    try:
        return yaml.safe_load(p.read_text(encoding="utf-8")) or {}
    except Exception:
        return None


def _resolve_template_deck_theme(template_meta: dict) -> dict:
    """Return the full deck theme for a template; fall back to a
    minimal synthesized dict from template_meta.theme_colors if the
    deck's theme.yaml isn't in the bundle (e.g. older builds)."""
    src = (template_meta.get("sources") or [{}])[0] if template_meta.get("sources") else {}
    deck_name = src.get("deck") or ""
    deck_theme = _load_deck_theme(deck_name)
    if deck_theme:
        return deck_theme
    palette: dict = {}
    aliases: dict = {}
    for role, hex_val in (template_meta.get("theme_colors") or {}).items():
        if hex_val:
            palette[role] = hex_val
            aliases[role] = role
    return {"palette": palette, "aliases": aliases}


def _build_scheme_remap(
    source_theme: dict | None, host_theme: dict | None
) -> dict[str, str]:
    """Build a clrScheme-slot remap aligning source.aliases with host.aliases.

    For each shared role where the slot differs between decks, add a
    rule. Keys are normalised (dk1/lt1/dk2/lt2) so both naming forms
    in the XML get matched.
    """
    if not source_theme or not host_theme:
        return {}
    src_aliases = source_theme.get("aliases") or {}
    host_aliases = host_theme.get("aliases") or {}
    remap: dict[str, str] = {}
    for role in ("primary", "accent", "text", "background"):
        src_slot = src_aliases.get(role)
        host_slot = host_aliases.get(role)
        if not src_slot or not host_slot:
            continue
        src_norm = _norm_scheme_slot(src_slot)
        host_norm = _norm_scheme_slot(host_slot)
        if src_norm != host_norm:
            remap[src_norm] = host_slot
    return remap


def _apply_scheme_remap(el, remap: dict[str, str]) -> int:
    """Rewrite <a:schemeClr val=...> per remap; returns count of substitutions."""
    if not remap:
        return 0
    try:
        from pptx.oxml.ns import qn
    except ImportError:
        return 0
    count = 0
    for sc in el.findall(".//" + qn("a:schemeClr")):
        val = (sc.get("val") or "").lower()
        norm = _norm_scheme_slot(val)
        if norm in remap:
            sc.set("val", remap[norm])
            count += 1
    return count


# ---------------------------------------------------------------------------
# v4.1 — surgical theme-font remap (D5 extension)
# ---------------------------------------------------------------------------
#
# Theme refs like <a:latin typeface="+mj-lt"/> already self-resolve to the
# host's major font after a cross-deck copy. Explicit typefaces don't —
# they survive the copy verbatim. Keynote-exported decks are the common
# offender: Keynote bakes "Helvetica Neue" in as an explicit typeface
# even though it IS that deck's theme major font. On a host whose theme
# major is e.g. Inter, the copied text then stubbornly renders Helvetica.
#
# Surgical policy: only rewrite an explicit typeface when it matches the
# *source* deck theme's major or minor. Anything else (Courier code,
# Comic Sans header) is preserved — assume the author meant it.

def _build_font_remap(
    source_theme: dict | None, host_theme: dict | None
) -> dict[str, str]:
    """Build a typeface remap aligning source major/minor to host's.

    Returns ``{lowercased_source_typeface: host_typeface}`` for the
    major and minor roles only when both sides have a font defined and
    they differ. One-off explicit fonts (i.e. typefaces that don't
    match the source theme's major/minor) are intentionally *not* in
    the remap and survive the copy unchanged.
    """
    if not source_theme or not host_theme:
        return {}
    src_fonts = source_theme.get("fonts") or {}
    host_fonts = host_theme.get("fonts") or {}
    remap: dict[str, str] = {}
    for role in ("major", "minor"):
        src = (src_fonts.get(role) or "").strip()
        host = (host_fonts.get(role) or "").strip()
        if not src or not host:
            continue
        if src.lower() == host.lower():
            continue
        remap[src.lower()] = host
    return remap


def _apply_font_remap(el, remap: dict[str, str]) -> int:
    """Rewrite ``<a:latin|ea|cs typeface="..."/>`` per remap; returns count.

    Only touches explicit typefaces. Theme refs (typefaces starting
    with ``+`` like ``+mj-lt``) self-resolve at render time and are
    skipped. Matching is case-insensitive; the substituted value
    preserves the host theme's casing.
    """
    if not remap:
        return 0
    try:
        from pptx.oxml.ns import qn
    except ImportError:
        return 0
    count = 0
    for tag in ("a:latin", "a:ea", "a:cs"):
        for node in el.findall(".//" + qn(tag)):
            typeface = (node.get("typeface") or "").strip()
            if not typeface or typeface.startswith("+"):
                continue
            host = remap.get(typeface.lower())
            if host and host != typeface:
                node.set("typeface", host)
                count += 1
    return count


# ---------------------------------------------------------------------------
# Rel-aware shape import (Phase 1 of the multi-slide compose fix)
# ---------------------------------------------------------------------------
#
# A naive `deepcopy(shape._element)` brings the source's rId references
# along verbatim. Those rIds live in the source slide's *local* rels file
# and mean nothing in the destination — PowerPoint then refuses to render
# the shape and pops the "couldn't read some content" dialog.
#
# _import_shape_xml handles this by:
#   1) deepcopy the shape XML,
#   2) for each rId-bearing attribute, look up the rel on the source part
#      and re-register it on the destination part (which auto-imports the
#      target part and any rels that come with it),
#   3) rewrite the attribute to the destination-local rId.
#
# Transitive rels (chart → embedded xlsx → image) ride along automatically
# because python-pptx packages parts together with their own rels graph
# when `relate_to(target_part, reltype)` registers a new internal rel.

_OPC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL_ATTRS = (
    f"{{{_OPC_REL_NS}}}embed",
    f"{{{_OPC_REL_NS}}}link",
    f"{{{_OPC_REL_NS}}}id",
)


def _iter_rel_attr_holders(el):
    """Yield (element, attr_qname) for every rId-bearing attribute under el."""
    for descendant in el.iter():
        for q in _REL_ATTRS:
            if q in descendant.attrib:
                yield descendant, q


def _existing_part_for_blob(package, partname, blob: bytes):
    """Return the dest-package part already at `partname` iff its blob
    matches `blob` (so the import can dedupe instead of colliding).
    None when no part exists there OR the existing part has different
    bytes — caller must then allocate a fresh partname."""
    for part in package.iter_parts():
        if part.partname == partname and getattr(part, "blob", None) == blob:
            return part
    return None


def _has_colliding_part(package, partname) -> bool:
    for part in package.iter_parts():
        if part.partname == partname:
            return True
    return False


def _clone_part_into(dest_package, src_part):
    """Materialise a copy of `src_part` inside `dest_package`. When the
    partname is free (or already holds the same bytes), reuse it as-is.
    When it would collide with different bytes, allocate a fresh
    partname so the OPC package stays valid.

    Returns a Part bound to dest_package."""
    from pptx.opc.package import Part
    from pptx.opc.packuri import PackURI

    src_partname = src_part.partname
    src_blob = src_part.blob
    src_ctype = src_part.content_type

    existing = _existing_part_for_blob(dest_package, src_partname, src_blob)
    if existing is not None:
        return existing

    if not _has_colliding_part(dest_package, src_partname):
        # Free slot — rebind by re-creating the part inside dest_package
        # using the same class so chart/image/etc. behaviours are kept.
        return type(src_part).load(src_partname, src_ctype, dest_package, src_blob)

    # Collision: allocate a fresh partname matching the original prefix
    # (e.g. /ppt/media/image%d.jpeg). PackURI exposes .baseURI + .ext.
    base_uri = src_partname.baseURI            # e.g. /ppt/media
    ext = src_partname.ext                     # e.g. jpeg
    # Strip trailing digits from the original filename to get the tmpl
    # stem ("image" from "image1"). Falls back to "part" if the source
    # used a non-numbered name.
    stem = src_partname.filename.split(".")[0]
    stem = re.sub(r"\d+$", "", stem) or "part"
    tmpl = f"{base_uri}/{stem}%d.{ext}" if ext else f"{base_uri}/{stem}%d"
    new_partname = dest_package.next_partname(tmpl)
    return type(src_part).load(
        PackURI(str(new_partname)), src_ctype, dest_package, src_blob,
    )


def _import_part_rel(src_part, dest_part, old_rid: str) -> str | None:
    """Look up old_rid on src_part, register the same target on dest_part,
    return the new dest-local rId. None if the rel cannot be resolved.

    Handles partname collisions across source decks by cloning the
    target onto dest_package with a fresh partname when needed."""
    try:
        src_rel = src_part.rels[old_rid]
    except KeyError:
        return None
    reltype = src_rel.reltype
    try:
        if src_rel.is_external:
            return dest_part.relate_to(src_rel.target_ref, reltype, is_external=True)
        target = src_rel.target_part
        dest_pkg = dest_part.package
        # If target was loaded from a different package OR its partname
        # would clash with an existing part in dest, clone into dest.
        if target.package is not dest_pkg or _has_colliding_part(
            dest_pkg, target.partname
        ):
            target = _clone_part_into(dest_pkg, target)
        return dest_part.relate_to(target, reltype)
    except Exception:
        # Unknown reltype or part class python-pptx can't handle — log
        # and skip so the rest of the shape import still succeeds.
        return None


def _rewrite_rel_attrs_in_subtree(src_part, dest_part, subtree) -> list[str]:
    """For every rId attr in `subtree`, import the rel onto dest_part and
    rewrite the attr value. Returns a list of warnings for refs that
    couldn't be resolved."""
    warnings: list[str] = []
    rid_subs: dict[str, str] = {}
    for el, attr_q in _iter_rel_attr_holders(subtree):
        old_rid = el.get(attr_q)
        if old_rid is None or old_rid in rid_subs:
            continue
        new_rid = _import_part_rel(src_part, dest_part, old_rid)
        if new_rid is None:
            warnings.append(
                f"unresolved rel {old_rid!r} on <{el.tag.rpartition('}')[2]} "
                f"{attr_q.rpartition('}')[2]}=...> — shape kept but ref dropped"
            )
            continue
        rid_subs[old_rid] = new_rid
    for el, attr_q in _iter_rel_attr_holders(subtree):
        old_rid = el.get(attr_q)
        if old_rid in rid_subs:
            el.set(attr_q, rid_subs[old_rid])
        elif old_rid is not None:
            # Drop the broken ref so OOXML doesn't see a phantom rId.
            del el.attrib[attr_q]
    return warnings


def _import_shape_xml(src_slide_part, dest_slide_part, dest_sptree, src_shape_el):
    """Import one shape: deepcopy, rewrite rIds against dest part, append.

    Returns a list of warnings (empty on the happy path)."""
    new_el = copy.deepcopy(src_shape_el)
    warnings = _rewrite_rel_attrs_in_subtree(src_slide_part, dest_slide_part, new_el)
    dest_sptree.append(new_el)
    return new_el, warnings


# ---------------------------------------------------------------------------
# Background flatten (Phase 3 of the multi-slide compose fix)
# ---------------------------------------------------------------------------
#
# PowerPoint resolves slide backgrounds through an inheritance chain:
#   slide → layout → master → theme (via <p:bgRef idx>)
# When we graft slides from one deck onto another, the destination's
# layout/master/theme is used — so the source's background is lost.
#
# _flatten_slide_background materialises the effective background as an
# explicit <p:bg> on the source slide before graft, so the background
# copies along with the shape tree.

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _qn(tag: str) -> str:
    prefix, _, local = tag.partition(":")
    ns = {"a": _A_NS, "p": _P_NS, "r": _OPC_REL_NS}.get(prefix)
    return f"{{{ns}}}{local}" if ns else tag


def _find_bg_element(cSld_el):
    if cSld_el is None:
        return None
    for child in cSld_el:
        if child.tag == _qn("p:bg"):
            return child
    return None


def _theme_part_for(slide_or_layout_or_master):
    """Return the SlideMaster theme part (the only theme that defines
    bgFillStyleLst entries used by <p:bgRef>)."""
    try:
        if hasattr(slide_or_layout_or_master, "slide_layout"):
            master = slide_or_layout_or_master.slide_layout.slide_master
        elif hasattr(slide_or_layout_or_master, "slide_master"):
            master = slide_or_layout_or_master.slide_master
        else:
            master = slide_or_layout_or_master
        # python-pptx exposes the theme part as master.element.part's
        # related theme part; easiest is to read its rels for the theme.
        master_part = master.part
        for rel in master_part.rels.values():
            if rel.reltype.endswith("/theme"):
                return rel.target_part
    except Exception:
        return None
    return None


def _resolve_theme_bg_ref(theme_part, idx: int):
    """Return the theme's bgFillStyleLst[idx-1001] element (deepcopy),
    suitable for embedding under <p:bgPr>. None on miss."""
    if theme_part is None or idx is None:
        return None
    try:
        from lxml import etree
        root = etree.fromstring(theme_part.blob)
    except Exception:
        return None
    bg_lst = root.find(
        f".//{_qn('a:fmtScheme')}/{_qn('a:bgFillStyleLst')}"
    )
    if bg_lst is None:
        return None
    children = list(bg_lst)
    # OOXML spec: bgRef@idx is 1001-based for bgFillStyleLst.
    pos = idx - 1001
    if pos < 0 or pos >= len(children):
        return None
    return copy.deepcopy(children[pos])


def _build_bg_pr_from_fill(fill_el):
    """Wrap a fill element (a:solidFill/a:gradFill/a:blipFill/a:pattFill)
    in a fresh <p:bgPr> ready to be the sole child of <p:bg>."""
    from lxml import etree
    bg_pr = etree.SubElement(etree.Element(_qn("p:bg")), _qn("p:bgPr"))
    bg_pr.append(fill_el)
    # OOXML schema: <p:bgPr> must end with <a:effectLst/> (can be empty).
    etree.SubElement(bg_pr, _qn("a:effectLst"))
    return bg_pr.getparent()  # the <p:bg> wrapper


def _find_effective_bg(slide):
    """Walk slide → layout → master and return the first <p:bg> element
    found (with the part it came from, for rel-import).

    Returns (bg_el_deepcopy, source_part) or (None, None)."""
    chain = []
    try:
        chain.append((slide._element.find(_qn("p:cSld")), slide.part))
    except Exception:
        pass
    try:
        layout = slide.slide_layout
        chain.append((layout._element.find(_qn("p:cSld")), layout.part))
    except Exception:
        pass
    try:
        master = slide.slide_layout.slide_master
        chain.append((master._element.find(_qn("p:cSld")), master.part))
    except Exception:
        pass
    for cSld_el, part in chain:
        bg = _find_bg_element(cSld_el)
        if bg is not None:
            return copy.deepcopy(bg), part
    return None, None


def _flatten_slide_background(slide) -> list[str]:
    """Materialise the slide's effective background into an explicit
    <p:bg> on the slide. Idempotent. Returns warnings list."""
    warnings: list[str] = []
    sld_el = slide._element
    cSld_el = sld_el.find(_qn("p:cSld"))
    if cSld_el is None:
        return warnings
    if _find_bg_element(cSld_el) is not None:
        return warnings  # already explicit

    bg_el, src_part = _find_effective_bg(slide)
    if bg_el is None or src_part is None:
        return warnings  # nothing to flatten

    # Resolve <p:bgRef> → inline <p:bgPr> from the theme.
    bg_ref = bg_el.find(_qn("p:bgRef"))
    if bg_ref is not None:
        try:
            idx = int(bg_ref.get("idx") or 0)
        except ValueError:
            idx = 0
        theme_part = _theme_part_for(slide)
        fill_el = _resolve_theme_bg_ref(theme_part, idx)
        if fill_el is None:
            warnings.append(
                f"bgRef idx={idx} could not be resolved against theme — "
                f"background may render blank"
            )
            return warnings
        # Carry over scheme-colour overrides from the bgRef (e.g. when
        # the slide says "use theme bg #2 but recolor accent1 → bg2").
        # bgRef wraps an inline colour spec as last child; preserve it.
        new_bg = _build_bg_pr_from_fill(fill_el)
        src_part = theme_part  # rels inside the theme blob, if any
    else:
        new_bg = bg_el

    # Import any rels referenced inside the bg subtree (e.g. blipFill
    # r:embed pointing at a media part on the source layout/master).
    warnings.extend(
        _rewrite_rel_attrs_in_subtree(src_part, slide.part, new_bg)
    )

    # <p:bg> must be the first child of <p:cSld> per the OOXML schema.
    cSld_el.insert(0, new_bg)
    return warnings


def _copy_slide_into(
    dest_prs: Presentation,
    src_slide_pptx: Path,
    source_theme: dict | None = None,
    host_theme: dict | None = None,
):
    """Copy the (single) slide from src_slide_pptx into dest_prs and return
    the new slide.

    Strategy: open the source deck, flatten its background so it's self-
    contained, copy its first slide's shape tree onto a blank layout in
    the destination via rel-aware import, then remap semantic clrScheme
    slots (D5) so the copied shapes stay consistent with the host's
    brand colours.
    """
    src_prs = Presentation(str(src_slide_pptx))
    if len(src_prs.slides) == 0:
        raise SystemExit(f"{src_slide_pptx}: no slides")
    src_slide = src_prs.slides[0]

    # Phase 3: make the source slide's background self-contained before
    # we graft. Otherwise the destination's blank layout supplies a
    # blank background and the original gradient/image is lost.
    _flatten_slide_background(src_slide)

    # Pick a blank-ish layout in dest. Prefer the layout with the fewest
    # placeholders to minimise interference with the copied content.
    dest_layout = None
    best_count = None
    for layout in dest_prs.slide_layouts:
        try:
            count = len(layout.placeholders)
        except Exception:
            count = 99
        if best_count is None or count < best_count:
            best_count = count
            dest_layout = layout
    if dest_layout is None:
        dest_layout = dest_prs.slide_layouts[0]

    new_slide = dest_prs.slides.add_slide(dest_layout)

    # Clear any placeholders the layout brought along — we want a clean canvas.
    for shp in list(new_slide.shapes):
        shp._element.getparent().remove(shp._element)

    # Carry the flattened background over to the new slide's <p:cSld>.
    src_cSld = src_slide._element.find(_qn("p:cSld"))
    dest_cSld = new_slide._element.find(_qn("p:cSld"))
    if src_cSld is not None and dest_cSld is not None:
        src_bg = _find_bg_element(src_cSld)
        if src_bg is not None and _find_bg_element(dest_cSld) is None:
            new_bg = copy.deepcopy(src_bg)
            # Rels inside the bg (blip fills) were imported onto
            # src_slide.part during flatten; re-import onto new_slide.part.
            _rewrite_rel_attrs_in_subtree(
                src_slide.part, new_slide.part, new_bg
            )
            dest_cSld.insert(0, new_bg)

    remap = _build_scheme_remap(source_theme, host_theme)
    font_remap = _build_font_remap(source_theme, host_theme)

    # Rel-aware shape import. Handles pictures, group shapes, hyperlinks,
    # picture-filled auto-shapes, charts, SmartArt — anything that carries
    # an rId reference inside its XML — by re-registering the underlying
    # parts on the destination slide and rewriting rIds in place.
    for shape in src_slide.shapes:
        new_el, _ws = _import_shape_xml(
            src_slide_part=src_slide.part,
            dest_slide_part=new_slide.part,
            dest_sptree=new_slide.shapes._spTree,
            src_shape_el=shape._element,
        )
        if remap:
            _apply_scheme_remap(new_el, remap)
        if font_remap:
            _apply_font_remap(new_el, font_remap)

    return new_slide


# ---------------------------------------------------------------------------
# Compose-mode (v4) — custom slides assembled from atoms + native shapes
# ---------------------------------------------------------------------------


def _frac_to_emu(spec: dict, key: str, total_emu: int, fallback: float) -> int:
    v = spec.get(key)
    if not isinstance(v, (int, float)) or v < 0:
        v = fallback
    return int(total_emu * float(v))


def _set_atom_geometry(el, x_emu: int, y_emu: int, w_emu: int, h_emu: int) -> None:
    """Rewrite (or insert) the offset+extent on a captured shape's XML.

    Tables and other GraphicFrame atoms use a direct ``p:xfrm`` child.
    Auto-shapes / freeforms (``p:sp``) use ``p:spPr/a:xfrm``. We handle
    both shapes silently — if neither path exists the atom keeps its
    source geometry rather than crashing.
    """
    try:
        from pptx.oxml.ns import qn
        from lxml import etree
    except ImportError:
        return
    xfrm = el.find(qn("p:xfrm"))
    if xfrm is None:
        sp_pr = el.find(qn("p:spPr"))
        if sp_pr is not None:
            xfrm = sp_pr.find(qn("a:xfrm"))
            if xfrm is None:
                xfrm = etree.SubElement(sp_pr, qn("a:xfrm"))
    if xfrm is None:
        return
    off = xfrm.find(qn("a:off"))
    if off is None:
        off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(x_emu))
    off.set("y", str(y_emu))
    ext = xfrm.find(qn("a:ext"))
    if ext is None:
        ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(w_emu))
    ext.set("cy", str(h_emu))


def _apply_recolor_xml(el, recolor: dict, deck_theme: dict | None = None) -> tuple[int, list[str]]:
    """Rewrite ``<a:srgbClr val="...">`` matches per a recolor map.

    Keys are source hex codes (with or without '#'). Values can be:
      - another hex code → direct substitution
      - a colour-role token ('primary', 'accent', 'text', 'background'
        or any clrScheme slot) → resolved against ``deck_theme`` if
        provided; warns and skips if unresolvable.
    Returns ``(substitutions_applied, warnings)``.
    """
    try:
        from pptx.oxml.ns import qn
    except ImportError:
        return 0, []
    warnings: list[str] = []
    palette = (deck_theme or {}).get("palette") or {}
    aliases = (deck_theme or {}).get("aliases") or {}

    def _resolve_target(raw: str) -> str | None:
        if not isinstance(raw, str):
            return None
        s = raw.strip().lstrip("#").upper()
        if len(s) == 6 and all(c in "0123456789ABCDEF" for c in s):
            return s
        # role / clrScheme token
        token = raw.strip().lower()
        alias_target = aliases.get(token)
        hex_val = palette.get(alias_target) if alias_target else palette.get(token)
        if not hex_val:
            return None
        return hex_val.lstrip("#").upper()

    norm_map: dict[str, str] = {}
    for src, tgt in recolor.items():
        if not isinstance(src, str):
            continue
        src_n = src.strip().lstrip("#").upper()
        if len(src_n) != 6:
            warnings.append(f"recolor source {src!r} is not a 6-digit hex; skipping")
            continue
        tgt_n = _resolve_target(tgt)
        if tgt_n is None:
            warnings.append(
                f"recolor target {tgt!r} could not be resolved (need hex or known role); skipping"
            )
            continue
        norm_map[src_n] = tgt_n

    if not norm_map:
        return 0, warnings

    applied = 0
    for clr in el.findall(".//" + qn("a:srgbClr")):
        val = (clr.get("val") or "").upper()
        if val in norm_map:
            clr.set("val", norm_map[val])
            applied += 1
    return applied, warnings


def _place_atom(
    slide,
    spec: dict,
    slide_w_emu: int,
    slide_h_emu: int,
    host_theme: dict | None = None,
) -> list[str]:
    """Place one atom on a slide at a fractional bbox.

    Supported spec keys:
      atom              asset id (required)
      x, y, w, h        fractions of slide; default to a top-left 40x30% box
      kind              optional override; defaults to the asset's stored kind
      recolor           {src_hex: target_hex_or_role} — rewrites srgbClr fills
      cells             list-of-lists (kind=table only) — post-place cell fill

    Picture atoms (raster/vector) re-import via ``add_picture`` so the
    image part lives in the host package. XML atoms (table, callout,
    freeform) graft as fragments. Charts / smartart are NOT yet
    placeable — they reference external parts that need a related-parts
    copy pass; we warn + skip rather than emit a half-broken slide.
    """
    warnings: list[str] = []
    aid = spec.get("atom")
    if not isinstance(aid, str) or not aid:
        warnings.append(f"compose-mode shape missing 'atom' id: {spec!r}")
        return warnings

    meta_p = asset_meta_path(aid)
    bin_p = asset_path(aid)
    if not meta_p.exists() or bin_p is None:
        warnings.append(f"compose-mode: asset {aid!r} not found in bundle")
        return warnings
    meta = yaml.safe_load(meta_p.read_text(encoding="utf-8")) or {}
    kind = spec.get("kind") or meta.get("kind") or ""

    x_emu = _frac_to_emu(spec, "x", slide_w_emu, 0.05)
    y_emu = _frac_to_emu(spec, "y", slide_h_emu, 0.05)
    w_emu = _frac_to_emu(spec, "w", slide_w_emu, 0.4)
    h_emu = _frac_to_emu(spec, "h", slide_h_emu, 0.3)

    if bin_p.suffix.lower() != ".xml":
        # Picture / vector — straightforward add_picture path.
        try:
            new_pic = slide.shapes.add_picture(
                str(bin_p), x_emu, y_emu, width=w_emu, height=h_emu
            )
            new_pic.name = aid
        except Exception as e:
            warnings.append(f"compose-mode: failed to place picture {aid!r}: {e}")
            return warnings
        if spec.get("recolor"):
            warnings.append(
                f"compose-mode {aid!r}: recolor on picture atoms not yet honored"
            )
        return warnings

    # XML atom path — graft as fragment.
    if kind in ("chart", "smartart"):
        warnings.append(
            f"compose-mode: {kind} atom {aid!r} not yet placeable "
            f"(related-parts copying deferred); skipping"
        )
        return warnings

    # Use pptx's parser so the element wraps with the right custom
    # class (CT_GraphicalObjectFrame, CT_Shape, …); a raw lxml element
    # would crash SlideShapeFactory's `has_ph_elm` check on lookup.
    try:
        from pptx.oxml import parse_xml
    except ImportError:
        warnings.append(f"compose-mode: pptx.oxml unavailable; cannot graft atom {aid!r}")
        return warnings

    try:
        new_el = parse_xml(bin_p.read_bytes())
    except Exception as e:
        warnings.append(f"compose-mode: atom {aid!r} XML failed to parse: {e}")
        return warnings

    _set_atom_geometry(new_el, x_emu, y_emu, w_emu, h_emu)

    # D5: remap semantic clrScheme slots when the atom came from a
    # different deck than the host — so e.g. an accent1 reference
    # painted in the source deck's accent still paints in the host
    # deck's accent on the rendered slide.
    asset_src = (meta.get("sources") or [{}])[0] if meta.get("sources") else {}
    asset_deck_theme = _load_deck_theme(asset_src.get("deck") or "")
    if asset_deck_theme and host_theme:
        scheme_remap = _build_scheme_remap(asset_deck_theme, host_theme)
        if scheme_remap:
            _apply_scheme_remap(new_el, scheme_remap)
        # v4.1: rewrite explicit typefaces that match the source's
        # major/minor → host's major/minor. Preserves intentional
        # one-off fonts (Courier, etc.).
        font_remap = _build_font_remap(asset_deck_theme, host_theme)
        if font_remap:
            _apply_font_remap(new_el, font_remap)

    recolor = spec.get("recolor")
    if isinstance(recolor, dict) and recolor:
        applied, ws = _apply_recolor_xml(new_el, recolor, host_theme)
        warnings.extend(ws)
        if applied == 0 and not ws:
            warnings.append(
                f"compose-mode: recolor for {aid!r} matched 0 colours in atom XML"
            )

    slide.shapes._spTree.append(new_el)

    if kind == "table" and spec.get("cells") is not None:
        new_shape = slide.shapes[-1]
        cells = spec.get("cells")
        if isinstance(cells, list) and all(isinstance(r, list) for r in cells):
            warnings.extend(_fill_table_shape(new_shape, cells))
        else:
            warnings.append(
                f"compose-mode: atom {aid!r} 'cells' must be list-of-lists; skipping fill"
            )

    return warnings


def _place_native_text(slide, spec: dict, slide_w_emu: int, slide_h_emu: int) -> list[str]:
    """Add a native textbox at fractional position from a compose-mode shape spec.

    Supported spec keys (in addition to x/y/w/h):
      value or text     text to render
      bold              bool
      font_role         informational — warns (not yet honored)
      color_role        informational — warns (not yet honored)
    """
    warnings: list[str] = []
    text = spec.get("value", spec.get("text", ""))
    if not isinstance(text, str):
        text = str(text)

    x_emu = _frac_to_emu(spec, "x", slide_w_emu, 0.05)
    y_emu = _frac_to_emu(spec, "y", slide_h_emu, 0.05)
    w_emu = _frac_to_emu(spec, "w", slide_w_emu, 0.5)
    h_emu = _frac_to_emu(spec, "h", slide_h_emu, 0.1)

    try:
        tb = slide.shapes.add_textbox(x_emu, y_emu, w_emu, h_emu)
    except Exception as e:
        warnings.append(f"compose-mode text: failed to add textbox: {e}")
        return warnings
    tb.text_frame.text = _strip_bullet_prefix(text)

    if spec.get("bold"):
        for p in tb.text_frame.paragraphs:
            for r in p.runs:
                r.font.bold = True

    for k in ("font_role", "color_role"):
        if spec.get(k):
            warnings.append(
                f"compose-mode text: {k}={spec[k]!r} not yet honored"
            )
    return warnings


def _compose_custom_slide(
    dest_prs: Presentation,
    entry: dict,
    host_theme: dict | None = None,
):
    """Append a blank slide and populate from ``entry.shapes``.

    Each shape spec is one of:
      - ``{"atom": <id>, ...}``  → routed to ``_place_atom``
      - ``{"kind": "text", "value": "...", ...}``  → routed to
        ``_place_native_text``

    Returns ``(new_slide, warnings)``.
    """
    warnings: list[str] = []

    dest_layout = None
    best_count = None
    for layout in dest_prs.slide_layouts:
        try:
            count = len(layout.placeholders)
        except Exception:
            count = 99
        if best_count is None or count < best_count:
            best_count = count
            dest_layout = layout
    if dest_layout is None:
        dest_layout = dest_prs.slide_layouts[0]

    new_slide = dest_prs.slides.add_slide(dest_layout)
    # Clean canvas — drop any layout-inherited placeholders.
    for shp in list(new_slide.shapes):
        shp._element.getparent().remove(shp._element)

    slide_w_emu = dest_prs.slide_width
    slide_h_emu = dest_prs.slide_height

    shapes = entry.get("shapes") or []
    for spec in shapes:
        if not isinstance(spec, dict):
            warnings.append(f"compose-mode entry: bad shape spec {spec!r}")
            continue
        if "atom" in spec:
            warnings.extend(_place_atom(new_slide, spec, slide_w_emu, slide_h_emu, host_theme))
        elif spec.get("kind") == "text":
            warnings.extend(_place_native_text(new_slide, spec, slide_w_emu, slide_h_emu))
        else:
            warnings.append(
                f"compose-mode entry: shape spec lacks 'atom' or kind=text: {spec!r}"
            )

    return new_slide, warnings


def _drop_first_slide(prs: Presentation) -> None:
    """Remove the host's original first slide after appending plan slides.

    Used when a compose-mode entry is the first plan entry — we still
    need a host pptx (for its theme/master), but we don't want its
    original first slide to leak into the output.
    """
    try:
        from pptx.oxml.ns import qn
    except ImportError:
        return
    sld_id_list = prs.slides._sldIdLst
    sld_ids = list(sld_id_list)
    if not sld_ids:
        return
    first = sld_ids[0]
    rid = first.get(qn("r:id"))
    sld_id_list.remove(first)
    if rid:
        try:
            prs.part.drop_rel(rid)
        except KeyError:
            pass


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------


def cmd_list(args: argparse.Namespace) -> None:
    index = load_index()
    flt = parse_filter(args.filter)
    out = {
        "templates": [t for t in index.get("templates", []) if matches_filter(t, flt)],
        "assets": [a for a in index.get("assets", []) if matches_filter(a, flt)],
    }
    json.dump(out, sys.stdout, indent=2, ensure_ascii=False)
    sys.stdout.write("\n")


def cmd_get(args: argparse.Namespace) -> None:
    target = args.id
    # Templates have a directory; assets have a sidecar.
    tdir = template_dir(target)
    if tdir.exists() and (tdir / "meta.yaml").exists():
        meta = yaml.safe_load((tdir / "meta.yaml").read_text(encoding="utf-8")) or {}
        out = {
            "kind": "template",
            "id": target,
            "meta": meta,
            "files": {
                "slide": str(tdir / "slide.pptx"),
                "meta": str(tdir / "meta.yaml"),
                "preview": str(tdir / "preview.png") if (tdir / "preview.png").exists() else None,
            },
        }
        json.dump(out, sys.stdout, indent=2, ensure_ascii=False)
        sys.stdout.write("\n")
        return

    apath = asset_meta_path(target)
    if apath.exists():
        meta = yaml.safe_load(apath.read_text(encoding="utf-8")) or {}
        bin_path = asset_path(target)
        out = {
            "kind": "asset",
            "id": target,
            "meta": meta,
            "files": {
                "binary": str(bin_path) if bin_path else None,
                "meta": str(apath),
            },
        }
        json.dump(out, sys.stdout, indent=2, ensure_ascii=False)
        sys.stdout.write("\n")
        return

    raise SystemExit(f"id not found: {target}")


def cmd_compose(args: argparse.Namespace) -> None:
    plan_path = Path(args.plan)
    out_path = Path(args.out)
    if not plan_path.exists():
        raise SystemExit(f"plan not found: {plan_path}")

    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    if not isinstance(plan, list) or not plan:
        raise SystemExit("plan must be a non-empty JSON array")

    plan_warnings: list[str] = []
    # Walk the plan once; each surviving entry is tagged 'template' or 'compose'.
    valid_entries: list[tuple[str, dict]] = []
    for entry in plan:
        if not isinstance(entry, dict):
            plan_warnings.append(f"skipping non-object plan entry: {entry!r}")
            continue
        if entry.get("compose"):
            valid_entries.append(("compose", entry))
            continue
        if "template" in entry:
            valid_entries.append(("template", entry))
            continue
        plan_warnings.append(
            f"skipping entry without 'template' or 'compose': {entry!r}"
        )

    if not valid_entries:
        raise SystemExit("plan has no valid entries (need 'template' or 'compose')")

    def template_meta(tid: str) -> dict:
        m = template_dir(tid) / "meta.yaml"
        if not m.exists():
            raise SystemExit(f"template not found: {tid}")
        return yaml.safe_load(m.read_text(encoding="utf-8")) or {}

    # Host pick: first template-mode entry wins. If the whole plan is
    # compose-mode, fall back to the first template alphabetically so
    # we still have a master/theme to render against.
    first_template_entry = next(
        (e for kind, e in valid_entries if kind == "template"), None
    )
    if first_template_entry is not None:
        host_tid = first_template_entry["template"]
        host_claims_first_slide = valid_entries[0][0] == "template" and (
            valid_entries[0][1]["template"] == host_tid
        )
    else:
        idx = load_index()
        templates = idx.get("templates") or []
        if not templates:
            raise SystemExit("no templates in bundle to serve as compose host")
        host_tid = sorted(t["id"] for t in templates)[0]
        host_claims_first_slide = False

    host_pptx = template_dir(host_tid) / "slide.pptx"
    if not host_pptx.exists():
        raise SystemExit(f"missing slide.pptx for {host_tid}")

    host_meta = template_meta(host_tid)
    host_theme = _resolve_template_deck_theme(host_meta)
    host_aspect = host_theme.get("aspect") or ""
    aspect_warned_for: set[str] = set()

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        shutil.copyfile(host_pptx, tmp.name)
        host_path = Path(tmp.name)

    try:
        dest_prs = Presentation(str(host_path))
        warnings: list[str] = list(plan_warnings)

        for i, (kind, entry) in enumerate(valid_entries):
            if kind == "template":
                tid = entry["template"]
                meta = template_meta(tid)
                src_theme = _resolve_template_deck_theme(meta)
                # Aspect-mismatch warning — once per foreign deck so
                # the agent sees it but the warning list stays terse.
                src_aspect = src_theme.get("aspect") or ""
                src_deck = ((meta.get("sources") or [{}])[0] or {}).get("deck") or ""
                if (
                    host_aspect and src_aspect and host_aspect != src_aspect
                    and src_deck and src_deck not in aspect_warned_for
                ):
                    warnings.append(
                        f"template {tid!r}: aspect {src_aspect!r} differs from "
                        f"host {host_aspect!r}; copied shapes are not auto-scaled"
                    )
                    aspect_warned_for.add(src_deck)
                if i == 0 and host_claims_first_slide:
                    slide = dest_prs.slides[0]
                    # Same flatten treatment as grafted slides so the
                    # host's first slide carries its background
                    # explicitly — otherwise inheritance loss in some
                    # decks would leave it blank while grafted slides
                    # render fine.
                    warnings.extend(_flatten_slide_background(slide))
                else:
                    src_pptx = template_dir(tid) / "slide.pptx"
                    if not src_pptx.exists():
                        raise SystemExit(f"missing slide.pptx for {tid}")
                    slide = _copy_slide_into(
                        dest_prs, src_pptx,
                        source_theme=src_theme, host_theme=host_theme,
                    )
                slots_by_id = {s["id"]: s for s in meta.get("slots", [])}
                for slot_id, value in (entry.get("slots") or {}).items():
                    kind_hint = slots_by_id.get(slot_id, {}).get("kind")
                    warnings.extend(_apply_slot_value(slide, slot_id, value, kind_hint))
            else:
                _, ws = _compose_custom_slide(dest_prs, entry, host_theme)
                warnings.extend(ws)

        # If the host's existing first slide wasn't claimed by an entry,
        # it's leftover scaffolding — drop it.
        if not host_claims_first_slide:
            _drop_first_slide(dest_prs)

        out_path.parent.mkdir(parents=True, exist_ok=True)
        dest_prs.save(str(out_path))
    finally:
        try:
            host_path.unlink()
        except OSError:
            pass

    result = {
        "output": str(out_path),
        "slides": len(valid_entries),
        "plan_entries": len(plan),
        "warnings": warnings,
    }
    json.dump(result, sys.stdout, indent=2, ensure_ascii=False)
    sys.stdout.write("\n")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(prog="reader.py")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_list = sub.add_parser("list", help="List templates + assets (optionally filtered).")
    p_list.add_argument("--filter", default=None, help="comma-separated key=value pairs")
    p_list.set_defaults(func=cmd_list)

    p_get = sub.add_parser("get", help="Get one template or asset by id.")
    p_get.add_argument("id")
    p_get.set_defaults(func=cmd_get)

    p_compose = sub.add_parser("compose", help="Compose a deck from a JSON plan.")
    p_compose.add_argument("plan")
    p_compose.add_argument("out")
    p_compose.set_defaults(func=cmd_compose)

    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
