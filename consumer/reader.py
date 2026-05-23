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


def _copy_slide_into(dest_prs: Presentation, src_slide_pptx: Path):
    """Copy the (single) slide from src_slide_pptx into dest_prs and return
    the new slide.

    Strategy: open the source deck, copy its first slide's shape tree onto a
    blank layout in the destination, then re-import image rels.
    """
    src_prs = Presentation(str(src_slide_pptx))
    if len(src_prs.slides) == 0:
        raise SystemExit(f"{src_slide_pptx}: no slides")
    src_slide = src_prs.slides[0]

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

    # Copy shapes from source. For pictures, re-add via add_picture so the
    # image part is imported into the destination package.
    for shape in src_slide.shapes:
        st = getattr(shape, "shape_type", None)
        # Picture: re-import bytes so the rel lives in the dest package.
        if st is not None and str(st).endswith("PICTURE"):
            try:
                blob = shape.image.blob
                ext = shape.image.ext or "png"
            except Exception:
                # Some pictures (e.g. background placeholders) may not expose .image
                blob = None
                ext = "png"
            if blob is not None:
                from io import BytesIO

                new_pic = new_slide.shapes.add_picture(
                    BytesIO(blob),
                    shape.left or 0,
                    shape.top or 0,
                    width=shape.width,
                    height=shape.height,
                )
                new_pic.name = shape.name
                continue
            # Fall through to plain XML copy if no blob accessible.

        el = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(el)

    return new_slide


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

    # v4: filter out compose-mode entries — full handling lands in
    # Phase D. Keep one warning per skipped entry so the agent sees
    # what wasn't honored. Template-mode entries pass through unchanged.
    plan_warnings: list[str] = []
    template_plan: list[dict] = []
    for entry in plan:
        if not isinstance(entry, dict):
            plan_warnings.append(f"skipping non-object plan entry: {entry!r}")
            continue
        if entry.get("compose"):
            plan_warnings.append(
                f"compose-mode entry skipped (Phase D will implement): "
                f"layout={entry.get('layout', '?')!r}, "
                f"shapes={len(entry.get('shapes') or [])}"
            )
            continue
        if "template" not in entry:
            plan_warnings.append(f"skipping entry with no `template` field: {entry!r}")
            continue
        template_plan.append(entry)

    if not template_plan:
        raise SystemExit(
            "plan has no template-mode entries to compose "
            "(compose-mode entries skipped pending Phase D)"
        )

    # Resolve template metadata for slot kind hints.
    def template_meta(tid: str) -> dict:
        m = template_dir(tid) / "meta.yaml"
        if not m.exists():
            raise SystemExit(f"template not found: {tid}")
        return yaml.safe_load(m.read_text(encoding="utf-8")) or {}

    # Start from the first template's deck as host so its master/theme applies.
    first_tid = template_plan[0]["template"]
    first_pptx = template_dir(first_tid) / "slide.pptx"
    if not first_pptx.exists():
        raise SystemExit(f"missing slide.pptx for {first_tid}")

    # Copy host into a tempfile so we don't mutate the bundle.
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        shutil.copyfile(first_pptx, tmp.name)
        host_path = Path(tmp.name)

    try:
        dest_prs = Presentation(str(host_path))

        warnings: list[str] = list(plan_warnings)

        # Fill slots on the host's existing first slide.
        first_meta = template_meta(first_tid)
        first_slots_by_id = {s["id"]: s for s in first_meta.get("slots", [])}
        first_slide = dest_prs.slides[0]
        for slot_id, value in (template_plan[0].get("slots") or {}).items():
            kind_hint = first_slots_by_id.get(slot_id, {}).get("kind")
            warnings.extend(_apply_slot_value(first_slide, slot_id, value, kind_hint))

        # Append subsequent slides.
        for entry in template_plan[1:]:
            tid = entry["template"]
            src_pptx = template_dir(tid) / "slide.pptx"
            if not src_pptx.exists():
                raise SystemExit(f"missing slide.pptx for {tid}")
            new_slide = _copy_slide_into(dest_prs, src_pptx)
            meta = template_meta(tid)
            slots_by_id = {s["id"]: s for s in meta.get("slots", [])}
            for slot_id, value in (entry.get("slots") or {}).items():
                kind_hint = slots_by_id.get(slot_id, {}).get("kind")
                warnings.extend(_apply_slot_value(new_slide, slot_id, value, kind_hint))

        out_path.parent.mkdir(parents=True, exist_ok=True)
        dest_prs.save(str(out_path))
    finally:
        try:
            host_path.unlink()
        except OSError:
            pass

    result = {
        "output": str(out_path),
        "slides": len(template_plan),
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
