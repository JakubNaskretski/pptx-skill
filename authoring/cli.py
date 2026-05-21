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
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Iterable

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
        "3:2": 3 / 2,
        "2:3": 2 / 3,
    }
    best, best_err = "free", 0.10  # tolerance
    for name, target in candidates.items():
        err = abs(r - target) / target
        if err < best_err:
            best, best_err = name, err
    return best


def position_quadrant(left: int, top: int, slide_w: int, slide_h: int) -> str:
    """Coarse position: left/center/right + top/middle/bottom."""
    cx = left + 0  # we use top-left for orientation
    cy = top + 0
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
        pos = position_quadrant(left, top, slide_w, slide_h)
        parts.append(f"{sid}@{pos}")
    return ", ".join(parts) if parts else "freeform"


def write_slide_fragment(src_path: Path, slide_idx: int, out_path: Path, renames: dict) -> None:
    """Save a single-slide fragment by reopening the deck and dropping others.

    The fragment keeps its slide masters/layouts so it can be re-composed.
    `renames` maps shape_id -> name we want to set so the consumer can
    locate slots at compose time.
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

    # Drop every other slide.
    sld_id_list = prs.slides._sldIdLst
    sld_ids = list(sld_id_list)
    for i, sld_id in enumerate(sld_ids):
        if i == slide_idx:
            continue
        rId = sld_id.rId
        sld_id_list.remove(sld_id)
        try:
            prs.part.drop_rel(rId)
        except Exception:
            pass

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))


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
        "feel": "",
        "composition": "",
        "colors": [],
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
SLIDE_FEEL_ENUM = {"formal", "punchy", "data-dense", "warm", "clinical", "celebratory"}
SLIDE_SUITABLE_ENUM = {
    "opener", "section_divider", "content", "data", "quote",
    "closing", "product", "team",
}

ASSET_REQUIRED = ("kind", "subject", "feel", "composition", "colors", "suitable_for")
ASSET_KIND_ENUM = {"photo", "icon", "logo", "illustration", "screenshot"}
ASSET_FEEL_ENUM = {"formal", "warm", "clinical", "punchy", "playful", "minimal", "dramatic"}
ASSET_COMPOSITION_ENUM = {
    "centered", "left-weighted", "right-weighted",
    "full-bleed", "top-heavy", "scattered",
}
ASSET_SUITABLE_ENUM = {
    "team", "hero", "product", "data", "culture", "event",
    "abstract", "decorative", "closing", "quote",
}


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


@cli.command()
def validate() -> None:
    """Schema-check every YAML; auto-promote complete ones to done."""
    failures: list[tuple[Path, list[str]]] = []
    promoted = 0
    total = 0

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


# --- preview ---------------------------------------------------------------


@cli.command()
def preview() -> None:
    """Best-effort PNG thumbnails of slide fragments via LibreOffice."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        click.echo("preview: LibreOffice not installed — skipping", err=True)
        return

    decks = WORKSPACE / "decks"
    if not decks.exists():
        click.echo("no decks ingested yet")
        return

    n_built = 0
    n_failed = 0
    for slide_pptx in sorted(decks.glob("*/slides/slide_*.pptx")):
        out_dir = slide_pptx.parent
        png_target = slide_pptx.with_suffix(".png")
        if png_target.exists() and png_target.stat().st_mtime >= slide_pptx.stat().st_mtime:
            continue
        # LibreOffice exits 0 even when it can't load the source. Verify
        # the output file actually appeared before counting it as built.
        try:
            subprocess.run(
                [
                    soffice, "--headless", "--convert-to", "png",
                    "--outdir", str(out_dir), str(slide_pptx),
                ],
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=120,
            )
        except (subprocess.SubprocessError, OSError) as e:
            click.echo(f"preview: failed for {slide_pptx.name}: {e}", err=True)
            n_failed += 1
            continue
        if png_target.exists():
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
                "feel": a.get("feel", ""),
                "composition": a.get("composition", ""),
                "colors": a.get("colors", []),
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


if __name__ == "__main__":
    cli()
