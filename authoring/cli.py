"""pptx-skill authoring CLI.

Turns example .pptx files into a portable template library. See PLAN.md
for the design. Seven commands: ingest, status, next, prompt, validate,
preview, build.

Workspace state lives under authoring/workspace/ and is gitignored.
Hand-edited YAML sidecars (slide_NN.yaml, <sha1>.yaml) are the unit of
description. `validate` auto-promotes complete pending sidecars to done.
`build` emits dist/skill.zip — the consumer artifact.
"""

from __future__ import annotations

import hashlib
import json
import os
import posixpath
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Iterable
from urllib.parse import unquote
from xml.etree import ElementTree as ET

import click
import yaml
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER


HERE = Path(__file__).resolve().parent
WORKSPACE = HERE / "workspace"
DIST = HERE / "dist"
SCHEMAS = HERE / "schemas"
PROMPTS = HERE / "prompts"
CONSUMER = HERE.parent / "consumer"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def ensure_workspace() -> None:
    (WORKSPACE / "decks").mkdir(parents=True, exist_ok=True)
    (WORKSPACE / "assets").mkdir(parents=True, exist_ok=True)


def read_yaml(path: Path) -> dict:
    text = path.read_text(encoding="utf-8")
    data = yaml.safe_load(text) or {}
    if not isinstance(data, dict):
        raise click.ClickException(f"{path}: expected mapping at top level")
    return data


def write_yaml(path: Path, data: dict) -> None:
    path.write_text(
        yaml.safe_dump(data, sort_keys=False, allow_unicode=True, width=100),
        encoding="utf-8",
    )


def aspect_ratio(width: int, height: int) -> str:
    """Pick a human aspect ratio token: 16:9, 4:3, 1:1, 3:4, 9:16, free."""
    if width <= 0 or height <= 0:
        return "free"
    r = width / height
    candidates = {
        "16:9": 16 / 9,
        "4:3": 4 / 3,
        "1:1": 1.0,
        "3:4": 3 / 4,
        "9:16": 9 / 16,
    }
    best, best_err = "free", 0.10  # tolerance
    for name, target in candidates.items():
        err = abs(r - target) / target
        if err < best_err:
            best, best_err = name, err
    return best


def position_quadrant(
    left: int, top: int, width: int, height: int, slide_w: int, slide_h: int
) -> str:
    """Coarse position: left/center/right + top/middle/bottom.

    Uses the shape's visual center, not its top-left, so a wide shape
    spanning the slide doesn't get mislabelled by its anchor corner.
    """
    cx = left + width // 2
    cy = top + height // 2
    if cx < slide_w / 3:
        h = "left"
    elif cx < 2 * slide_w / 3:
        h = "center"
    else:
        h = "right"
    if cy < slide_h / 3:
        v = "top"
    elif cy < 2 * slide_h / 3:
        v = "middle"
    else:
        v = "bottom"
    return f"{v}-{h}"


# ---------------------------------------------------------------------------
# Ingest
# ---------------------------------------------------------------------------


PLACEHOLDER_TEXTUAL = {
    PP_PLACEHOLDER.TITLE,
    PP_PLACEHOLDER.SUBTITLE,
    PP_PLACEHOLDER.BODY,
    PP_PLACEHOLDER.CENTER_TITLE,
    PP_PLACEHOLDER.OBJECT,
}


def _shape_is_picture(shape) -> bool:
    return getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE


def _slot_id_for_placeholder(ph_type, used: set) -> str:
    if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
        base = "title"
    elif ph_type == PP_PLACEHOLDER.SUBTITLE:
        base = "subtitle"
    elif ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT):
        base = "body"
    elif ph_type == PP_PLACEHOLDER.PICTURE:
        base = "hero"
    else:
        base = "field"
    return _unique_id(base, used)


def _unique_id(base: str, used: set) -> str:
    if base not in used:
        used.add(base)
        return base
    i = 2
    while f"{base}_{i}" in used:
        i += 1
    out = f"{base}_{i}"
    used.add(out)
    return out


def detect_slots(slide, slide_w: int, slide_h: int) -> tuple[list[dict], dict]:
    """Detect slot definitions on a slide.

    Returns (slots, shape_renames) where shape_renames maps the original
    shape element id to the slot id we want set as shape.name (so the
    consumer can find each slot by shape name at compose time).
    """
    slots: list[dict] = []
    used_ids: set = set()
    renames: dict = {}  # shape_id (int) -> slot_id

    slide_area = slide_w * slide_h

    # First pass: placeholders.
    for shape in list(slide.placeholders):
        ph = shape.placeholder_format
        ph_type = ph.type
        if ph_type == PP_PLACEHOLDER.PICTURE:
            slot_id = _unique_id("hero", used_ids)
            slots.append(
                {
                    "id": slot_id,
                    "kind": "image",
                    "aspect": aspect_ratio(shape.width or 0, shape.height or 0),
                }
            )
            renames[shape.shape_id] = slot_id
            continue
        if ph_type not in PLACEHOLDER_TEXTUAL:
            # Decorative / unsupported — leave frozen.
            continue
        # Textual placeholder.
        tf = getattr(shape, "text_frame", None)
        if tf is None:
            continue
        text = tf.text or ""
        paragraphs = [p for p in tf.paragraphs]
        is_bulleted = (
            ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)
            and len(paragraphs) > 1
            and any((p.text or "").strip() for p in paragraphs)
        )
        slot_id = _slot_id_for_placeholder(ph_type, used_ids)
        if is_bulleted:
            non_empty = [p for p in paragraphs if (p.text or "").strip()]
            slots.append(
                {
                    "id": slot_id,
                    "kind": "bullets",
                    "max_items": max(1, len(non_empty)),
                }
            )
        else:
            slots.append(
                {
                    "id": slot_id,
                    "kind": "text",
                    "max_chars": max(20, int(len(text) * 1.5) or 60),
                }
            )
        renames[shape.shape_id] = slot_id

    # Second pass: free pictures > 20% slide area.
    for shape in list(slide.shapes):
        if not _shape_is_picture(shape):
            continue
        if getattr(shape, "is_placeholder", False):
            continue
        w = shape.width or 0
        h = shape.height or 0
        if slide_area <= 0:
            continue
        if (w * h) / slide_area <= 0.20:
            continue  # frozen background / logo / decoration
        slot_id = _unique_id("hero", used_ids)
        slots.append(
            {
                "id": slot_id,
                "kind": "image",
                "aspect": aspect_ratio(w, h),
            }
        )
        renames[shape.shape_id] = slot_id

    return slots, renames


def infer_layout(slide, slots: list[dict], renames: dict, slide_w: int, slide_h: int) -> str:
    """Best-effort one-line layout description from slot positions."""
    parts: list[str] = []
    for shape in list(slide.shapes):
        sid = renames.get(shape.shape_id)
        if not sid:
            continue
        left = shape.left or 0
        top = shape.top or 0
        width = shape.width or 0
        height = shape.height or 0
        pos = position_quadrant(left, top, width, height, slide_w, slide_h)
        parts.append(f"{sid}@{pos}")
    return ", ".join(parts) if parts else "freeform"


def write_slide_fragment(src_path: Path, slide_idx: int, out_path: Path, renames: dict) -> None:
    """Save a single-slide fragment by reopening the deck and dropping others.

    The fragment keeps its slide masters/layouts so it can be re-composed.
    `renames` maps shape_id -> name we want to set so the consumer can
    locate slots at compose time.

    After python-pptx writes the package, we garbage-collect orphan
    parts (other slides' media, layouts no slide references, etc.) so
    each fragment carries only what its slide actually needs.
    """
    prs = Presentation(str(src_path))

    # Apply renames first (on the current slide, before we drop the others).
    if 0 <= slide_idx < len(prs.slides):
        slide = prs.slides[slide_idx]
        for shape in list(slide.shapes):
            if shape.shape_id in renames:
                try:
                    shape.name = renames[shape.shape_id]
                except Exception:
                    pass

    # Drop every other slide from sldIdLst + the package relationship.
    # KeyError is the only python-pptx-documented failure for drop_rel
    # (rId already gone); anything else is real and should surface.
    sld_id_list = prs.slides._sldIdLst
    sld_ids = list(sld_id_list)
    for i, sld_id in enumerate(sld_ids):
        if i == slide_idx:
            continue
        rId = sld_id.rId
        sld_id_list.remove(sld_id)
        try:
            prs.part.drop_rel(rId)
        except KeyError:
            pass

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))

    # Prune orphans introduced by python-pptx's "leave parts behind" model.
    prune_unreachable_parts(out_path)


# --- Fragment GC: keep only parts reachable from /_rels/.rels ---------------
#
# python-pptx's save() prunes unused layout *rels* when it writes a
# fragment (e.g. all the layouts other slides used) but leaves orphan
# `<p:sldLayoutId>` entries inside the master XML. Those orphans then
# point to rIds that no longer exist — and python-pptx (and PowerPoint)
# blow up when iterating the master's layouts. It also leaves any other
# orphan parts (notes themes, dropped images, etc.) inside the zip.
#
# This pass:
#  1. For each master, removes `<p:sldLayoutId>` entries whose r:id is
#     no longer in the master's rels file.
#  2. Walks /_rels/.rels through every part's rels, collecting the
#     closure of reachable parts.
#  3. Filters `[Content_Types].xml` Override entries to match.
#  4. Rewrites the zip with only reachable parts.

_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _rels_path_for(part_path: str) -> str:
    dir_part, _, name = part_path.rpartition("/")
    return f"{dir_part}/_rels/{name}.rels" if dir_part else f"_rels/{name}.rels"


def _resolve_rel(rels_path: str, target: str) -> str:
    target = unquote(target)
    if rels_path == "_rels/.rels":
        base = ""
    else:
        base = posixpath.dirname(posixpath.dirname(rels_path))
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base, target)) if base else target


def _parse_rels(xml_bytes: bytes) -> list[tuple[str, str]]:
    """Return list of (target, type) for non-External relationships."""
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return []
    out: list[tuple[str, str]] = []
    for rel in root.findall(f"{{{_REL_NS}}}Relationship"):
        if rel.get("TargetMode") == "External":
            continue
        out.append((rel.get("Target", ""), rel.get("Type", "")))
    return out


_PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
# `r:` inside document XML (slideMaster.xml, presentation.xml, etc.)
# binds to officeDocument/relationships — distinct from the package
# /relationships namespace used by .rels files themselves.
_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _valid_rids(rels_xml: bytes) -> set[str]:
    """Return the Id set declared by a rels file."""
    try:
        root = ET.fromstring(rels_xml)
    except ET.ParseError:
        return set()
    return {rel.get("Id", "") for rel in root.findall(f"{{{_REL_NS}}}Relationship")}


def _layouts_used_by_slides(zf: zipfile.ZipFile, names: list[str]) -> set[str]:
    """Layouts referenced via rels by any slide still in the package."""
    used: set[str] = set()
    names_set = set(names)
    for name in names:
        if not (name.startswith("ppt/slides/slide") and name.endswith(".xml")):
            continue
        rels_path = _rels_path_for(name)
        if rels_path not in names_set:
            continue
        for target, rtype in _parse_rels(zf.read(rels_path)):
            if rtype.endswith("/slideLayout"):
                used.add(_resolve_rel(rels_path, target))
    return used


def _prune_master_rels(rels_xml: bytes, rels_path: str, kept_layouts: set[str]) -> bytes:
    """Drop slideLayout rels not in kept_layouts. Returns rels_xml unchanged
    if nothing to drop (identity-comparable)."""
    try:
        root = ET.fromstring(rels_xml)
    except ET.ParseError:
        return rels_xml
    changed = False
    for rel in list(root.findall(f"{{{_REL_NS}}}Relationship")):
        if not rel.get("Type", "").endswith("/slideLayout"):
            continue
        if _resolve_rel(rels_path, rel.get("Target", "")) not in kept_layouts:
            root.remove(rel)
            changed = True
    if not changed:
        return rels_xml
    ET.register_namespace("", _REL_NS)
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + \
        ET.tostring(root, encoding="utf-8")


def _repair_master_xml(master_xml: bytes, valid_rids: set[str]) -> bytes:
    """Drop <p:sldLayoutId> entries whose r:id is no longer in the master's
    rels. python-pptx trims unused layouts from the rels file when it
    saves a fragment, but leaves orphan sldLayoutId entries in the master
    XML — which then break loading because each sldLayoutId is resolved
    via its rId."""
    try:
        root = ET.fromstring(master_xml)
    except ET.ParseError:
        return master_xml
    rid_attr = f"{{{_DOC_REL_NS}}}id"
    changed = False
    for lst in root.findall(f"{{{_PML_NS}}}sldLayoutIdLst"):
        for entry in list(lst.findall(f"{{{_PML_NS}}}sldLayoutId")):
            if entry.get(rid_attr) not in valid_rids:
                lst.remove(entry)
                changed = True
    if not changed:
        return master_xml
    ET.register_namespace("p", _PML_NS)
    ET.register_namespace("a", _DML_NS)
    ET.register_namespace("r", _DOC_REL_NS)
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + \
        ET.tostring(root, encoding="utf-8")


def _master_xml_path_for_rels(rels_path: str) -> str:
    # 'ppt/slideMasters/_rels/slideMaster1.xml.rels'
    # -> 'ppt/slideMasters/slideMaster1.xml'
    parent = posixpath.dirname(posixpath.dirname(rels_path))
    name = posixpath.basename(rels_path)
    if name.endswith(".rels"):
        name = name[: -len(".rels")]
    return f"{parent}/{name}"


def prune_unreachable_parts(pptx_path: Path) -> tuple[int, int]:
    """Rewrite a .pptx in place, keeping only parts reachable from the
    package root rels. Returns (kept, removed) part counts."""
    with zipfile.ZipFile(pptx_path, "r") as zf:
        all_names = list(zf.namelist())
        name_set = set(all_names)

        # Active pruning + orphan repair for each master.
        # 1. Drop slideLayout rels not used by any surviving slide.
        # 2. Then drop <p:sldLayoutId> entries in the master XML pointing
        #    to rIds that no longer exist (covers both our prune AND any
        #    orphans python-pptx left behind from its own save logic).
        kept_layouts = _layouts_used_by_slides(zf, all_names)
        rewritten: dict[str, bytes] = {}
        for name in all_names:
            if "/slideMasters/_rels/" not in name or not name.endswith(".rels"):
                continue
            master_xml_path = _master_xml_path_for_rels(name)
            if master_xml_path not in name_set:
                continue

            original_rels = zf.read(name)
            new_rels = _prune_master_rels(original_rels, name, kept_layouts)
            if new_rels is not original_rels:
                rewritten[name] = new_rels

            valid = _valid_rids(new_rels)
            original_master = zf.read(master_xml_path)
            repaired = _repair_master_xml(original_master, valid)
            if repaired is not original_master:
                rewritten[master_xml_path] = repaired

        reachable: set[str] = set()
        for keep in ("[Content_Types].xml", "_rels/.rels"):
            if keep in name_set:
                reachable.add(keep)

        queue: list[str] = ["_rels/.rels"]
        seen_rels: set[str] = set()
        while queue:
            rels_path = queue.pop(0)
            if rels_path in seen_rels or rels_path not in name_set:
                continue
            seen_rels.add(rels_path)
            reachable.add(rels_path)
            xml = rewritten.get(rels_path, zf.read(rels_path))
            for target, _ in _parse_rels(xml):
                resolved = _resolve_rel(rels_path, target)
                if resolved in name_set and resolved not in reachable:
                    reachable.add(resolved)
                    part_rels = _rels_path_for(resolved)
                    if part_rels in name_set and part_rels not in seen_rels:
                        queue.append(part_rels)

        try:
            ct_root = ET.fromstring(zf.read("[Content_Types].xml"))
            for override in list(ct_root.findall(f"{{{_CT_NS}}}Override")):
                pn = override.get("PartName", "").lstrip("/")
                if pn not in reachable:
                    ct_root.remove(override)
            ET.register_namespace("", _CT_NS)
            new_ct = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + \
                ET.tostring(ct_root, encoding="utf-8")
        except ET.ParseError:
            new_ct = zf.read("[Content_Types].xml")

        tmp_path = pptx_path.with_suffix(pptx_path.suffix + ".pruning")
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for name in all_names:
                if name not in reachable:
                    continue
                if name == "[Content_Types].xml":
                    zf_out.writestr(name, new_ct)
                elif name in rewritten:
                    zf_out.writestr(name, rewritten[name])
                else:
                    zf_out.writestr(name, zf.read(name))

    tmp_path.replace(pptx_path)
    return len(reachable), len(name_set) - len(reachable)


def write_slide_yaml_stub(out_path: Path, slide_id: str, layout: str, slots: list[dict], deck_stem: str, slide_number: int) -> None:
    """Write a slide sidecar stub. Skips writing if file exists with status: done/locked."""
    if out_path.exists():
        existing = read_yaml(out_path)
        status = existing.get("status", "pending")
        if status in ("done", "locked"):
            # Refresh structural fields only, leave descriptive fields alone.
            existing["id"] = slide_id
            existing["layout"] = layout
            existing["slots"] = slots
            existing["sources"] = [{"deck": deck_stem, "slide": slide_number}]
            write_yaml(out_path, existing)
            return

    stub = {
        "intent": "",
        "feel": "",
        "suitable_for": [],
        "status": "pending",
        "notes": "",
        "id": slide_id,
        "layout": layout,
        "slots": slots,
        "sources": [{"deck": deck_stem, "slide": slide_number}],
    }
    write_yaml(out_path, stub)


def write_asset_yaml_stub(out_path: Path, asset_id: str, sha1: str, deck_stem: str, slide_number: int) -> None:
    if out_path.exists():
        existing = read_yaml(out_path)
        # Append source if new.
        sources = existing.get("sources") or []
        if not any(s.get("deck") == deck_stem and s.get("slide") == slide_number for s in sources):
            sources.append({"deck": deck_stem, "slide": slide_number})
            existing["sources"] = sources
        existing["id"] = asset_id
        existing["sha1"] = sha1
        write_yaml(out_path, existing)
        return

    stub = {
        "kind": "",
        "subject": "",
        "depicts": "",
        "feel": "",
        "composition": "",
        "colors": [],
        "scope": [],
        "suitable_for": [],
        "status": "pending",
        "notes": "",
        "id": asset_id,
        "sha1": sha1,
        "sources": [{"deck": deck_stem, "slide": slide_number}],
    }
    write_yaml(out_path, stub)


def extract_picture_assets(slide, deck_stem: str, slide_number: int, assets_dir: Path) -> list[str]:
    """Save all picture shapes on this slide as <sha1>.<ext> binaries + yaml stubs.

    Returns the list of asset ids extracted (sha1 prefixes).
    """
    extracted: list[str] = []
    for shape in list(slide.shapes):
        if not _shape_is_picture(shape):
            continue
        try:
            image = shape.image
        except Exception:
            continue
        blob = image.blob
        ext = (image.ext or "png").lstrip(".")
        sha = hashlib.sha1(blob).hexdigest()
        asset_id = f"asset_{sha[:8]}"
        bin_path = assets_dir / f"{sha}.{ext}"
        yaml_path = assets_dir / f"{sha}.yaml"
        if not bin_path.exists():
            bin_path.write_bytes(blob)
        write_asset_yaml_stub(yaml_path, asset_id, sha, deck_stem, slide_number)
        extracted.append(asset_id)
    return extracted


# ---------------------------------------------------------------------------
# Validate
# ---------------------------------------------------------------------------


SLIDE_REQUIRED = ("intent", "feel", "suitable_for")
ASSET_REQUIRED = ("kind", "subject", "feel", "composition", "colors", "scope", "suitable_for")

VOCAB_PATH = SCHEMAS / "vocab.yaml"


def _load_vocab() -> dict:
    with VOCAB_PATH.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f)


VOCAB = _load_vocab()
SLIDE_FEEL_ENUM = set(VOCAB["slide"]["feel"])
SLIDE_SUITABLE_ENUM = set(VOCAB["slide"]["suitable_for"])
ASSET_KIND_ENUM = set(VOCAB["asset"]["kind"])
ASSET_FEEL_ENUM = set(VOCAB["asset"]["feel"])
ASSET_COMPOSITION_ENUM = set(VOCAB["asset"]["composition"])
ASSET_SUITABLE_ENUM = set(VOCAB["asset"]["suitable_for"])
ASSET_SCOPE_PREFIXES = set(VOCAB["asset"]["scope_prefixes"])


def _missing_or_empty(data: dict, keys: Iterable[str]) -> list[str]:
    missing = []
    for k in keys:
        v = data.get(k)
        if v is None or v == "" or v == []:
            missing.append(k)
    return missing


def validate_slide(data: dict) -> list[str]:
    errors = _missing_or_empty(data, SLIDE_REQUIRED)
    feel = data.get("feel")
    if feel and feel not in SLIDE_FEEL_ENUM:
        errors.append(f"feel '{feel}' not in {sorted(SLIDE_FEEL_ENUM)}")
    sfor = data.get("suitable_for") or []
    if isinstance(sfor, list):
        bad = [t for t in sfor if t not in SLIDE_SUITABLE_ENUM]
        if bad:
            errors.append(f"suitable_for has unknown tag(s): {bad}")
    intent = data.get("intent") or ""
    if intent and len(intent.split()) > 25:
        errors.append("intent is over 25 words; prefer <=20")
    return errors


def validate_asset(data: dict) -> list[str]:
    errors = _missing_or_empty(data, ASSET_REQUIRED)
    if data.get("kind") and data["kind"] not in ASSET_KIND_ENUM:
        errors.append(f"kind '{data['kind']}' not in {sorted(ASSET_KIND_ENUM)}")
    if data.get("feel") and data["feel"] not in ASSET_FEEL_ENUM:
        errors.append(f"feel '{data['feel']}' not in {sorted(ASSET_FEEL_ENUM)}")
    if data.get("composition") and data["composition"] not in ASSET_COMPOSITION_ENUM:
        errors.append(
            f"composition '{data['composition']}' not in {sorted(ASSET_COMPOSITION_ENUM)}"
        )
    sfor = data.get("suitable_for") or []
    if isinstance(sfor, list):
        bad = [t for t in sfor if t not in ASSET_SUITABLE_ENUM]
        if bad:
            errors.append(f"suitable_for has unknown tag(s): {bad}")
    scope = data.get("scope") or []
    if isinstance(scope, list):
        for entry in scope:
            if entry == "generic":
                continue
            if not isinstance(entry, str) or ":" not in entry:
                errors.append(
                    f"scope entry {entry!r} must be 'generic' or '<prefix>:<value>'"
                )
                continue
            prefix, _, value = entry.partition(":")
            if prefix not in ASSET_SCOPE_PREFIXES:
                errors.append(
                    f"scope entry {entry!r}: unknown prefix '{prefix}' "
                    f"(use one of {sorted(ASSET_SCOPE_PREFIXES)} or 'generic')"
                )
            if not value:
                errors.append(f"scope entry {entry!r}: missing value after ':'")
    subj = data.get("subject") or ""
    if subj and len(subj.split()) > 30:
        errors.append("subject is over 30 words; prefer <=25")
    return errors


# ---------------------------------------------------------------------------
# Discovery helpers
# ---------------------------------------------------------------------------


def iter_slide_yamls() -> Iterable[Path]:
    decks = WORKSPACE / "decks"
    if not decks.exists():
        return []
    return sorted(decks.glob("*/slides/slide_*.yaml"))


def iter_asset_yamls() -> Iterable[Path]:
    assets = WORKSPACE / "assets"
    if not assets.exists():
        return []
    return sorted(assets.glob("*.yaml"))


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


@click.group()
def cli() -> None:
    """pptx-skill authoring CLI."""


# --- ingest ----------------------------------------------------------------


@cli.command()
@click.argument("deck", type=click.Path(exists=True, dir_okay=False, path_type=Path))
def ingest(deck: Path) -> None:
    """Strip a deck into workspace/ as slide fragments + asset binaries."""
    ensure_workspace()
    deck_stem = deck.stem
    deck_dir = WORKSPACE / "decks" / deck_stem
    slides_dir = deck_dir / "slides"
    slides_dir.mkdir(parents=True, exist_ok=True)

    # Snapshot the original so future ingests are reproducible.
    original = deck_dir / "original.pptx"
    if not original.exists() or original.read_bytes() != deck.read_bytes():
        shutil.copyfile(deck, original)

    prs = Presentation(str(original))
    slide_w = prs.slide_width or 9144000
    slide_h = prs.slide_height or 6858000

    assets_dir = WORKSPACE / "assets"

    n_slides = len(prs.slides)
    for idx in range(n_slides):
        slide = prs.slides[idx]
        slot_defs, renames = detect_slots(slide, slide_w, slide_h)
        layout = infer_layout(slide, slot_defs, renames, slide_w, slide_h)

        slide_number = idx + 1
        slide_id = f"{deck_stem}_{slide_number:02d}"
        slide_pptx = slides_dir / f"slide_{slide_number:02d}.pptx"
        slide_yaml = slides_dir / f"slide_{slide_number:02d}.yaml"

        write_slide_fragment(original, idx, slide_pptx, renames)
        write_slide_yaml_stub(slide_yaml, slide_id, layout, slot_defs, deck_stem, slide_number)
        extract_picture_assets(slide, deck_stem, slide_number, assets_dir)

    click.echo(f"Ingested {deck.name}: {n_slides} slide(s) → {slides_dir}")


# --- status ----------------------------------------------------------------


def _bucket(paths: Iterable[Path]) -> tuple[list[Path], list[Path], list[Path]]:
    pending, done, locked = [], [], []
    for p in paths:
        try:
            data = read_yaml(p)
        except Exception:
            pending.append(p)
            continue
        status = data.get("status", "pending")
        if status == "done":
            done.append(p)
        elif status == "locked":
            locked.append(p)
        else:
            pending.append(p)
    return pending, done, locked


@cli.command()
def status() -> None:
    """Counts pending/done/locked; lists pending file paths."""
    slides = list(iter_slide_yamls())
    assets = list(iter_asset_yamls())
    sp, sd, sl = _bucket(slides)
    ap, ad, al = _bucket(assets)
    click.echo(f"slides:  pending={len(sp)} done={len(sd)} locked={len(sl)} total={len(slides)}")
    click.echo(f"assets:  pending={len(ap)} done={len(ad)} locked={len(al)} total={len(assets)}")
    if sp or ap:
        click.echo("\npending:")
        for p in sp + ap:
            click.echo(f"  {p.relative_to(HERE)}")


# --- next ------------------------------------------------------------------


@cli.command(name="next")
@click.option("--kind", type=click.Choice(["asset", "slide"]))
@click.option("--open", "open_editor", is_flag=True, help="Launch $EDITOR on the YAML.")
def next_cmd(kind, open_editor: bool) -> None:
    """Print the next pending item path."""
    if kind == "asset":
        candidates = list(iter_asset_yamls())
    elif kind == "slide":
        candidates = list(iter_slide_yamls())
    else:
        candidates = list(iter_slide_yamls()) + list(iter_asset_yamls())

    target: Path | None = None
    for p in candidates:
        try:
            data = read_yaml(p)
        except Exception:
            target = p
            break
        if data.get("status", "pending") == "pending":
            target = p
            break

    if target is None:
        click.echo("nothing pending")
        return

    click.echo(str(target))

    if open_editor:
        editor = os.environ.get("EDITOR", "vi")
        try:
            subprocess.run([editor, str(target)], check=False)
        except FileNotFoundError:
            click.echo(f"warn: $EDITOR not found ({editor})", err=True)


# --- prompt ----------------------------------------------------------------


@cli.command()
@click.option("--kind", type=click.Choice(["asset", "slide"]), default="asset")
def prompt(kind: str) -> None:
    """Print the bundled describe prompt to stdout for copy/paste."""
    path = PROMPTS / f"describe_{kind}.md"
    if not path.exists():
        raise click.ClickException(f"missing prompt file: {path}")
    click.echo(path.read_text(encoding="utf-8"))


# --- validate --------------------------------------------------------------


def check_prompt_drift() -> list[str]:
    """Warn if a vocab enum value is missing from its prompt markdown."""
    pairs = [
        (PROMPTS / "describe_slide.md", "slide.feel", VOCAB["slide"]["feel"]),
        (PROMPTS / "describe_slide.md", "slide.suitable_for", VOCAB["slide"]["suitable_for"]),
        (PROMPTS / "describe_asset.md", "asset.kind", VOCAB["asset"]["kind"]),
        (PROMPTS / "describe_asset.md", "asset.feel", VOCAB["asset"]["feel"]),
        (PROMPTS / "describe_asset.md", "asset.composition", VOCAB["asset"]["composition"]),
        (PROMPTS / "describe_asset.md", "asset.suitable_for", VOCAB["asset"]["suitable_for"]),
    ]
    errs: list[str] = []
    for path, label, values in pairs:
        if not path.exists():
            errs.append(f"{path.name}: missing prompt file")
            continue
        text = path.read_text(encoding="utf-8")
        missing = [v for v in values if v not in text]
        if missing:
            errs.append(
                f"{path.name}: {label} missing value(s) {missing} "
                f"(update prompt to match schemas/vocab.yaml)"
            )
    return errs


@cli.command()
def validate() -> None:
    """Schema-check every YAML; auto-promote complete ones to done."""
    failures: list[tuple[Path, list[str]]] = []
    promoted = 0
    total = 0

    drift = check_prompt_drift()
    if drift:
        failures.append((VOCAB_PATH, drift))

    for p in iter_slide_yamls():
        total += 1
        try:
            data = read_yaml(p)
        except Exception as e:
            failures.append((p, [f"unreadable: {e}"]))
            continue
        errs = validate_slide(data)
        if errs:
            failures.append((p, errs))
            continue
        if data.get("status") == "pending":
            data["status"] = "done"
            write_yaml(p, data)
            promoted += 1

    for p in iter_asset_yamls():
        total += 1
        try:
            data = read_yaml(p)
        except Exception as e:
            failures.append((p, [f"unreadable: {e}"]))
            continue
        errs = validate_asset(data)
        if errs:
            failures.append((p, errs))
            continue
        if data.get("status") == "pending":
            data["status"] = "done"
            write_yaml(p, data)
            promoted += 1

    if failures:
        click.echo(f"FAIL: {len(failures)} sidecar(s) have errors")
        for p, errs in failures:
            click.echo(f"  {p.relative_to(HERE)}")
            for e in errs:
                click.echo(f"    - {e}")
    click.echo(f"checked {total}, promoted pending→done: {promoted}")
    if failures:
        sys.exit(1)


# --- slide rendering -------------------------------------------------------
#
# Slide fragments are rendered to PNGs by the first available backend.
# Priority: PowerPoint COM (Windows, best fidelity if installed) →
# LibreOffice headless (any OS) → macOS Quick Look. Both `cli.py preview`
# and the bulk-batch flow in `app.py` go through `render_slide_to_png`.


def _render_via_powerpoint_com(slide_pptx: Path, out_png: Path, size: int) -> bool:
    if sys.platform != "win32":
        return False
    ps = shutil.which("powershell") or shutil.which("pwsh")
    if not ps:
        return False
    # PowerShell accepts forward slashes on Windows, sidesteps quoting.
    src = str(slide_pptx.resolve()).replace("\\", "/")
    dst = str(out_png.resolve()).replace("\\", "/")
    height = int(size * 9 / 16)  # assume 16:9; aspect drift fine for thumbnails
    script = (
        '$ErrorActionPreference = "Stop"; '
        '$pp = New-Object -ComObject PowerPoint.Application; '
        f'$deck = $pp.Presentations.Open("{src}", $true, $true, $false); '
        f'$deck.Slides.Item(1).Export("{dst}", "PNG", {size}, {height}); '
        '$deck.Close(); $pp.Quit();'
    )
    try:
        subprocess.run(
            [ps, "-NoProfile", "-NonInteractive", "-Command", script],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=60,
        )
    except (subprocess.SubprocessError, OSError):
        return False
    return out_png.exists()


def _render_via_libreoffice(slide_pptx: Path, out_png: Path, size: int) -> bool:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    out_dir = out_png.parent
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "png",
             "--outdir", str(out_dir), str(slide_pptx)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=120,
        )
    except (subprocess.SubprocessError, OSError):
        return False
    produced = out_dir / f"{slide_pptx.stem}.png"
    if produced != out_png and produced.exists():
        produced.replace(out_png)
    return out_png.exists()


def _render_via_qlmanage(slide_pptx: Path, out_png: Path, size: int) -> bool:
    if sys.platform != "darwin":
        return False
    ql = shutil.which("qlmanage")
    if not ql:
        return False
    out_dir = out_png.parent
    try:
        subprocess.run(
            [ql, "-t", "-s", str(size), "-o", str(out_dir), str(slide_pptx)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=20,
        )
    except (subprocess.SubprocessError, OSError):
        return False
    # qlmanage writes <name>.pptx.png; normalise to <stem>.png.
    produced = out_dir / f"{slide_pptx.name}.png"
    if produced != out_png and produced.exists():
        produced.replace(out_png)
    return out_png.exists()


_RENDERERS = (
    ("PowerPoint COM (Windows)", _render_via_powerpoint_com),
    ("LibreOffice (soffice)", _render_via_libreoffice),
    ("macOS Quick Look (qlmanage)", _render_via_qlmanage),
)


def render_slide_to_png(
    slide_pptx: Path, out_png: Path | None = None, size: int = 1200
) -> Path | None:
    """Render a single-slide .pptx to a PNG. Returns the PNG path on
    success, None if no backend could produce one. Cached by mtime."""
    if out_png is None:
        out_png = slide_pptx.with_suffix(".png")
    if out_png.exists() and out_png.stat().st_mtime >= slide_pptx.stat().st_mtime:
        return out_png
    for _name, fn in _RENDERERS:
        if fn(slide_pptx, out_png, size):
            return out_png
    return None


def available_renderers() -> list[str]:
    """Human-readable list of slide-rendering backends that *look* usable
    (binaries on PATH; doesn't verify PowerPoint is actually installed)."""
    out: list[str] = []
    if sys.platform == "win32" and (shutil.which("powershell") or shutil.which("pwsh")):
        out.append("PowerPoint COM (Windows; requires PowerPoint installed)")
    if shutil.which("soffice") or shutil.which("libreoffice"):
        out.append("LibreOffice (soffice)")
    if sys.platform == "darwin" and shutil.which("qlmanage"):
        out.append("macOS Quick Look (qlmanage)")
    return out


# --- preview ---------------------------------------------------------------


@cli.command()
def preview() -> None:
    """Best-effort PNG thumbnails of slide fragments. Tries PowerPoint
    (Windows), LibreOffice, then macOS Quick Look in order."""
    if not available_renderers():
        click.echo(
            "preview: no slide renderer available — install LibreOffice, "
            "or use PowerPoint on Windows",
            err=True,
        )
        return

    decks = WORKSPACE / "decks"
    if not decks.exists():
        click.echo("no decks ingested yet")
        return

    n_built = 0
    n_failed = 0
    for slide_pptx in sorted(decks.glob("*/slides/slide_*.pptx")):
        png = render_slide_to_png(slide_pptx)
        if png is not None:
            n_built += 1
        else:
            n_failed += 1
    if n_failed:
        click.echo(f"preview: {n_failed} slide(s) could not be rendered", err=True)
    click.echo(f"preview: built {n_built} thumbnail(s)")


# --- build -----------------------------------------------------------------


def _ext_for(blob_path: Path) -> str:
    return blob_path.suffix.lstrip(".") or "bin"


def build_index(slides: list[dict], assets: list[dict]) -> dict:
    return {
        "templates": [
            {
                "id": s["id"],
                "intent": s.get("intent", ""),
                "feel": s.get("feel", ""),
                "suitable_for": s.get("suitable_for", []),
                "layout": s.get("layout", ""),
                "slots": s.get("slots", []),
            }
            for s in slides
        ],
        "assets": [
            {
                "id": a["id"],
                "kind": a.get("kind", ""),
                "subject": a.get("subject", ""),
                "depicts": a.get("depicts", ""),
                "feel": a.get("feel", ""),
                "composition": a.get("composition", ""),
                "colors": a.get("colors", []),
                "scope": a.get("scope", []),
                "suitable_for": a.get("suitable_for", []),
            }
            for a in assets
        ],
    }


@cli.command()
@click.option("--allow-pending", is_flag=True, help="Build even if some sidecars are pending.")
@click.option("--out", "out_path", type=click.Path(path_type=Path), default=None)
def build(allow_pending: bool, out_path: Path | None) -> None:
    """Emit dist/skill.zip — the consumer artifact."""
    slide_yamls = list(iter_slide_yamls())
    asset_yamls = list(iter_asset_yamls())

    if not slide_yamls and not asset_yamls:
        raise click.ClickException("workspace is empty; run `ingest` first")

    pending = []
    slides_meta: list[dict] = []
    assets_meta: list[dict] = []

    for p in slide_yamls:
        d = read_yaml(p)
        if d.get("status", "pending") == "pending":
            pending.append(p)
        d["_yaml_path"] = p
        slides_meta.append(d)

    for p in asset_yamls:
        d = read_yaml(p)
        if d.get("status", "pending") == "pending":
            pending.append(p)
        d["_yaml_path"] = p
        assets_meta.append(d)

    if pending and not allow_pending:
        click.echo(f"refusing to build: {len(pending)} sidecar(s) still pending", err=True)
        for p in pending:
            click.echo(f"  {p.relative_to(HERE)}", err=True)
        click.echo("re-run with --allow-pending to override", err=True)
        sys.exit(1)

    DIST.mkdir(parents=True, exist_ok=True)
    zip_path = out_path or (DIST / "skill.zip")

    reader_src = CONSUMER / "reader.py"
    skill_md = CONSUMER / "SKILL.md"
    consumer_reqs = CONSUMER / "requirements.txt"
    brand_md = HERE / "brand.md"
    if not reader_src.exists() or not skill_md.exists():
        raise click.ClickException(
            "consumer/reader.py or SKILL.md missing — cannot build zip"
        )

    index = build_index(slides_meta, assets_meta)

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("SKILL.md", skill_md.read_text(encoding="utf-8"))
        zf.writestr("reader.py", reader_src.read_text(encoding="utf-8"))
        if consumer_reqs.exists():
            zf.writestr("requirements.txt", consumer_reqs.read_text(encoding="utf-8"))
        zf.writestr("index.json", json.dumps(index, indent=2, ensure_ascii=False))
        if brand_md.exists():
            brand_text = brand_md.read_text(encoding="utf-8").strip()
            if brand_text:
                zf.writestr("brand.md", brand_text + "\n")

        for sd in slides_meta:
            tid = sd["id"]
            yaml_path: Path = sd["_yaml_path"]
            slide_pptx = yaml_path.with_suffix(".pptx")
            if not slide_pptx.exists():
                click.echo(f"warn: missing {slide_pptx} — skipping {tid}", err=True)
                continue
            zf.write(slide_pptx, f"templates/{tid}/slide.pptx")
            clean = {k: v for k, v in sd.items() if not k.startswith("_")}
            zf.writestr(
                f"templates/{tid}/meta.yaml",
                yaml.safe_dump(clean, sort_keys=False, allow_unicode=True),
            )
            preview_png = yaml_path.with_suffix(".png")
            if preview_png.exists():
                zf.write(preview_png, f"templates/{tid}/preview.png")

        for ad in assets_meta:
            aid = ad["id"]
            sha = ad.get("sha1", "")
            yaml_path: Path = ad["_yaml_path"]
            # Find the binary alongside the yaml — same stem, any extension.
            blob: Path | None = None
            for cand in yaml_path.parent.glob(f"{yaml_path.stem}.*"):
                if cand.suffix == ".yaml":
                    continue
                blob = cand
                break
            if blob is None:
                click.echo(f"warn: no binary for asset {aid} ({sha}) — skipping", err=True)
                continue
            ext = _ext_for(blob)
            zf.write(blob, f"assets/{aid}.{ext}")
            clean = {k: v for k, v in ad.items() if not k.startswith("_")}
            zf.writestr(
                f"assets/{aid}.yaml",
                yaml.safe_dump(clean, sort_keys=False, allow_unicode=True),
            )

    click.echo(
        f"built {zip_path} — {len(slides_meta)} template(s), {len(assets_meta)} asset(s)"
    )


# --- package-app ----------------------------------------------------------
#
# Produces a zip containing ONLY the application code + scaffolding —
# nothing organisation-specific. Use it to hand the tool to a teammate
# without leaking KB content, built artifacts, source decks, or
# org-tuned brand rules. Allowlist-based: anything not explicitly named
# below is excluded.

# Files included verbatim (path is relative to repo root). Globs allowed.
PACKAGE_APP_ALLOWLIST = (
    "README.md",
    "PLAN.md",
    ".gitignore",
    "authoring/cli.py",
    "authoring/app.py",
    "authoring/requirements.txt",
    "authoring/prompts/*.md",
    "authoring/schemas/*.yaml",
    "authoring/schemas/*.yml",
    "consumer/SKILL.md",
    "consumer/reader.py",
    "consumer/requirements.txt",
    "consumer/index.example.json",
)


@cli.command(name="package-app")
@click.option("--out", "out_path", type=click.Path(path_type=Path), default=None,
              help="Output zip path. Default: authoring/dist/pptx-skill-app.zip")
def package_app(out_path: Path | None) -> None:
    """Build a transferable zip of the app code only.

    Never includes workspace/, dist/, source .pptx files, brand.md, or
    other session/local-only files. Ships `brand.example.md` so the
    recipient knows where to put their own brand rules.
    """
    REPO_ROOT = HERE.parent
    out_path = out_path or (DIST / "pptx-skill-app.zip")
    DIST.mkdir(parents=True, exist_ok=True)

    seen: set[str] = set()
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for pattern in PACKAGE_APP_ALLOWLIST:
            matches = sorted(REPO_ROOT.glob(pattern))
            if not matches:
                continue
            for src in matches:
                if not src.is_file():
                    continue
                arc = src.relative_to(REPO_ROOT).as_posix()
                if arc in seen:
                    continue
                seen.add(arc)
                zf.write(src, arc)

        # Brand stub goes in as brand.example.md so the recipient can copy
        # it to brand.md and fill in. Always written from a clean template,
        # never from the local (possibly org-customised) brand.md.
        zf.writestr(
            "authoring/brand.example.md",
            _BRAND_STUB,
        )

        # Empty workspace dir + a README pointer for first-time users.
        zf.writestr("authoring/workspace/.gitkeep", "")
        zf.writestr("FIRST_RUN.md", _FIRST_RUN_NOTES)

    size_kb = out_path.stat().st_size // 1024
    click.echo(f"packaged {out_path} — {len(seen) + 3} file(s), {size_kb} KB")


_BRAND_STUB = """# Brand & visual rules

Copy this file to `authoring/brand.md` and fill in. The compose page
auto-includes brand.md in every prompt bundle so the LLM sees these
rules before reading the brief.

Keep it short and prescriptive. Long preambles get ignored.

## Palette

- Primary: <e.g. #0A2540 navy>
- Accent:  <e.g. #F26B38 orange>
- Avoid:   <colors that look off-brand>

## Voice

- <e.g. "factual, no hype, third person">

## Don'ts

- <e.g. "no exclamation marks in headlines">
- <e.g. "never use stock people-in-suits imagery">
"""


_FIRST_RUN_NOTES = """# First run

This zip contains the pptx-skill app code only. No KB content, no
built artifacts. Steps to get going on a fresh machine:

1. Unzip somewhere on disk.
2. `pip install -r authoring/requirements.txt`
3. `cp authoring/brand.example.md authoring/brand.md` and edit to your
   org's palette / voice / taboos.
4. Drop a source deck somewhere and `python3 authoring/cli.py ingest path/to/deck.pptx`.
5. Describe slides/assets via `python3 authoring/app.py` (compose page
   served on http://127.0.0.1:5050) or hand-edit YAML in
   `authoring/workspace/`.
6. `python3 authoring/cli.py validate` and then `... build` to produce
   the durable consumer skill.zip, or use the compose page to generate
   per-deck agent prompt bundles.

See README.md and PLAN.md for the design.
"""


if __name__ == "__main__":
    cli()
