"""pptx-skill describe app — local Flask UI for filling YAML sidecars.

Run:
    python3 authoring/app.py
Opens http://127.0.0.1:5000 in the default browser.

Authoring-only; not shipped to the consumer. Talks to the same YAML files
the CLI uses. Slide previews are generated lazily via macOS `qlmanage`.
"""

from __future__ import annotations

import io
import json as json_mod
import os
import shutil
import subprocess
import sys
import tempfile
import webbrowser
import zipfile
from datetime import datetime
from pathlib import Path
from threading import Timer

import yaml
from flask import Flask, abort, jsonify, render_template_string, request, send_file

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import cli as cli_mod  # noqa: E402

WORKSPACE = HERE / "workspace"
BATCHES_DIR = WORKSPACE / "_batches"

# Per-request user-supplied assets (images / svg / xml the user attaches
# to a specific compose request). Separate from the workspace KB.
#   _user_assets/staged/  — accumulates uploads until Download bundle
#   _user_assets/bundle/  — snapshot of the last bundle's user assets,
#                           used by compose-run to resolve user_<id>
#                           references the agent emits in the plan.
USER_ASSETS_DIR = WORKSPACE / "_user_assets"
USER_STAGED_DIR = USER_ASSETS_DIR / "staged"
USER_BUNDLE_DIR = USER_ASSETS_DIR / "bundle"

app = Flask(__name__)
# Cap uploads at 200 MB — typical decks are <50 MB; this leaves headroom
# without letting a stray multi-GB upload exhaust /tmp.
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _safe_read(p: Path) -> dict:
    try:
        return cli_mod.read_yaml(p)
    except Exception:
        return {}


def _yaml_for_rel(rel: str) -> Path:
    candidate = (HERE / rel).resolve()
    try:
        candidate.relative_to(WORKSPACE.resolve())
    except ValueError:
        abort(400, "path outside workspace")
    if not candidate.exists():
        abort(404, "yaml not found")
    return candidate


def _kind_of(yaml_path: Path) -> str:
    return "slide" if yaml_path.parent.name == "slides" else "asset"


def _items() -> dict:
    out: dict = {"slides": [], "assets": []}
    for p in cli_mod.iter_slide_yamls():
        d = _safe_read(p)
        out["slides"].append({
            "id": d.get("id") or p.stem,
            "yaml": str(p.relative_to(HERE)),
            "status": d.get("status", "pending"),
        })
    for p in cli_mod.iter_asset_yamls():
        d = _safe_read(p)
        out["assets"].append({
            "id": d.get("id") or p.stem,
            "yaml": str(p.relative_to(HERE)),
            "status": d.get("status", "pending"),
        })
    return out


def _ensure_slide_png(slide_pptx: Path) -> Path | None:
    return cli_mod.render_slide_to_png(slide_pptx)


def _asset_binary(yaml_path: Path) -> Path | None:
    for cand in yaml_path.parent.glob(f"{yaml_path.stem}.*"):
        if cand.suffix != ".yaml":
            return cand
    return None


def _assets_for_slide(slide_yaml_path: Path) -> list[Path]:
    """Return asset yaml paths whose `sources` includes this slide.

    Each asset.yaml carries a `sources: [{deck, slide}, ...]` list that
    ingest populates. We reverse-lookup: scan the workspace's asset pool
    and pick the ones that mention this deck+slide pair. Used by the
    slide_with_assets bulk-describe mode so each slide bundle ships
    with the actual binaries of every picture / atom on it.
    """
    # Slide yamls live at workspace/decks/<deck>/slides/slide_NN.yaml — pull
    # the deck stem from the path and the slide number from the filename.
    try:
        deck_stem = slide_yaml_path.parent.parent.name
        slide_number = int(slide_yaml_path.stem.removeprefix("slide_"))
    except (ValueError, AttributeError):
        return []
    out: list[Path] = []
    for ap in cli_mod.iter_asset_yamls():
        try:
            data = cli_mod.read_yaml(ap)
        except Exception:
            continue
        for src in (data.get("sources") or []):
            if (
                isinstance(src, dict)
                and src.get("deck") == deck_stem
                and src.get("slide") == slide_number
            ):
                out.append(ap)
                break
    return out


SLIDE_DESCRIPTIVE = ("intent", "feel", "suitable_for", "notes", "interpretation")
ASSET_DESCRIPTIVE = (
    "kind", "subject", "depicts", "feel", "composition",
    "colors", "scope", "suitable_for", "notes", "interpretation",
)
_LIST_KEYS = {"suitable_for", "colors", "scope"}


def _descriptive_yaml(data: dict, kind: str) -> str:
    keys = SLIDE_DESCRIPTIVE if kind == "slide" else ASSET_DESCRIPTIVE
    subset: dict = {}
    for k in keys:
        if k in data:
            subset[k] = data[k]
        elif k in _LIST_KEYS:
            subset[k] = []
        else:
            subset[k] = ""
    return yaml.safe_dump(subset, sort_keys=False, allow_unicode=True, width=100)


def _strip_yaml_fence(txt: str) -> str:
    """Strip ```yaml / ```json / bare ``` code fences from LLM output."""
    txt = txt.strip()
    if txt.startswith("```yaml"):
        txt = txt[len("```yaml"):].lstrip()
    elif txt.startswith("```json"):
        txt = txt[len("```json"):].lstrip()
    elif txt.startswith("```"):
        txt = txt[3:].lstrip()
    if txt.endswith("```"):
        txt = txt[:-3].rstrip()
    return txt


def _save_and_validate(p: Path, merge_fields: dict) -> tuple[list, str]:
    existing = _safe_read(p)
    existing.update(merge_fields)
    cli_mod.write_yaml(p, existing)
    kind = _kind_of(p)
    errs = (
        cli_mod.validate_slide(existing)
        if kind == "slide"
        else cli_mod.validate_asset(existing)
    )
    if not errs and existing.get("status") != "locked":
        existing["status"] = "done"
        cli_mod.write_yaml(p, existing)
    return errs, existing.get("status", "pending")


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------


@app.get("/api/items")
def api_items():
    return jsonify(_items())


def _safe_pptx_filename(raw: str) -> str | None:
    """Sanitize a user-uploaded filename to a safe basename ending in .pptx.

    Strips any path components and leading dots so a malicious upload
    can't escape /tmp. Preserves unicode filenames as-is.
    Returns None for empties, non-.pptx, or whitespace-only names.
    """
    if not raw:
        return None
    base = Path(raw).name.lstrip(".").strip()
    if not base or not base.lower().endswith(".pptx"):
        return None
    return base


@app.post("/api/ingest")
def api_ingest():
    """Upload + ingest a .pptx file. Rejects with 409 if the deck stem
    already exists under workspace/decks/ (delete it first to re-ingest).

    Form field: ``pptx`` (multipart file).
    Returns: ``{"deck_stem", "slides", "pictures", "atoms"}`` on success.
    """
    if "pptx" not in request.files:
        return jsonify({"error": "no file uploaded (expect form field 'pptx')"}), 400
    f = request.files["pptx"]
    safe_name = _safe_pptx_filename(f.filename or "")
    if safe_name is None:
        return jsonify({"error": "filename must end in .pptx"}), 400

    # Save under a temp dir so the original filename is preserved
    # verbatim (cli_mod._ingest_pptx uses path.stem for the deck name).
    with tempfile.TemporaryDirectory(prefix="pptx_upload_") as td:
        tmp_path = Path(td) / safe_name
        f.save(str(tmp_path))
        try:
            result = cli_mod._ingest_pptx(tmp_path, reject_collision=True)
        except cli_mod.IngestCollisionError as e:
            return jsonify({
                "error": (
                    f"deck '{e.deck_stem}' is already ingested. Delete "
                    f"workspace/decks/{e.deck_stem}/ before re-uploading."
                ),
                "deck_stem": e.deck_stem,
            }), 409
        except Exception as e:
            return jsonify({"error": f"ingest failed: {e}"}), 500
    return jsonify(result)


@app.get("/api/item")
def api_item():
    rel = request.args.get("yaml", "")
    p = _yaml_for_rel(rel)
    kind = _kind_of(p)
    data = _safe_read(p)
    return jsonify({
        "kind": kind,
        "yaml": rel,
        "data": data,
        "yaml_text": _descriptive_yaml(data, kind),
    })


@app.post("/api/save")
def api_save():
    body = request.get_json(force=True) or {}
    rel = body.get("yaml", "")
    fields = body.get("fields", {}) or {}
    p = _yaml_for_rel(rel)
    errs, status = _save_and_validate(p, fields)
    return jsonify({"errors": errs, "status": status})


@app.post("/api/save-raw")
def api_save_raw():
    body = request.get_json(force=True) or {}
    rel = body.get("yaml", "")
    raw = body.get("text", "")
    txt = _strip_yaml_fence(raw)
    try:
        parsed = yaml.safe_load(txt) if txt else {}
    except Exception as e:
        return jsonify({"errors": [f"YAML parse error: {e}"], "status": "pending"})
    if parsed is None:
        parsed = {}
    if not isinstance(parsed, dict):
        return jsonify({
            "errors": ["expected a YAML mapping at top level"], "status": "pending",
        })
    p = _yaml_for_rel(rel)
    errs, status = _save_and_validate(p, parsed)
    return jsonify({"errors": errs, "status": status})


@app.post("/api/parse-yaml")
def api_parse_yaml():
    raw = (request.get_json(force=True) or {}).get("text", "")
    txt = _strip_yaml_fence(raw)
    try:
        parsed = yaml.safe_load(txt)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    if not isinstance(parsed, dict):
        return jsonify({"error": "expected a YAML mapping at top level"}), 400
    return jsonify({"data": parsed})


@app.get("/api/prompt")
def api_prompt():
    kind = request.args.get("kind", "asset")
    if kind not in ("asset", "slide"):
        abort(400)
    text = (HERE / "prompts" / f"describe_{kind}.md").read_text(encoding="utf-8")
    return jsonify({"text": text})


# ---------------------------------------------------------------------------
# Batch describe
# ---------------------------------------------------------------------------


def _bulk_instructions_slide_with_assets(
    n: int,
    slide_prompt: str,
    asset_prompt: str,
    items: dict,
) -> str:
    """Bundled-describe prompt: each numbered folder = one slide + its assets.

    `items` maps "01" → {"slide": rel, "assets": {asset_id: rel}} so the
    prompt can show the model the exact asset_id keys it should use in
    each entry's `assets` block (rather than expecting the model to
    invent or copy them from filenames).
    """
    # Render a small table of expected asset keys per bundle so the model
    # knows exactly what to fill in. Keeps the JSON output structured.
    expected_lines = []
    for key in sorted(items.keys()):
        bundle = items[key]
        asset_ids = list((bundle.get("assets") or {}).keys())
        if asset_ids:
            ids_str = ", ".join(f"`{a}`" for a in asset_ids)
            expected_lines.append(f"- **{key}/** — slide + {len(asset_ids)} asset(s): {ids_str}")
        else:
            expected_lines.append(f"- **{key}/** — slide + 0 assets (describe slide only)")
    expected_block = "\n".join(expected_lines)

    sample_block = (
        '{\n'
        '  "01": {\n'
        '    "slide": {\n'
        '      "intent": "...",\n'
        '      "feel": "formal",\n'
        '      "suitable_for": ["opener"],\n'
        '      "notes": "",\n'
        '      "interpretation": ""\n'
        '    },\n'
        '    "assets": {\n'
        '      "asset_abc12345": {\n'
        '        "kind": "photo",\n'
        '        "subject": "...",\n'
        '        "depicts": "...",\n'
        '        "feel": "warm",\n'
        '        "composition": "centered",\n'
        '        "colors": ["navy"],\n'
        '        "scope": ["generic"],\n'
        '        "suitable_for": ["team"],\n'
        '        "notes": "",\n'
        '        "interpretation": ""\n'
        '      }\n'
        '    }\n'
        '  },\n'
        '  "02": { "slide": {...}, "assets": { ... } }\n'
        '}\n'
    )
    return (
        f"# Bulk describe batch — {n} slide(s) with their constituent assets\n\n"
        f"The downstream pipeline picks slides as templates and pulls "
        f"individual assets (photos, logos, tables, callouts) onto them at "
        f"compose time. Both need descriptions. Today they're described "
        f"independently — this batch lets you describe them **together**, "
        f"so the slide's context can inform asset descriptions (and vice "
        f"versa).\n\n"
        f"## Bundle structure\n\n"
        f"The zip contains {n} numbered folders. Each folder is one slide:\n\n"
        f"- `<NN>/slide.png` — rendered preview of the slide\n"
        f"- `<NN>/<asset_id>.<ext>` — the binary of every asset that "
        f"appears on this slide (photo, logo, icon, table xml, etc.). "
        f"The filename stem is the asset's stable id — use it verbatim "
        f"as the JSON key.\n\n"
        f"### Expected per-bundle contents\n\n"
        f"{expected_block}\n\n"
        f"## Output format\n\n"
        f"Return ONE JSON object. Top-level keys are the 2-digit bundle "
        f"ids (`\"01\"`, `\"02\"`, ..., `\"{n:02d}\"`). Each value is an "
        f"object with two sub-keys:\n\n"
        f"- `slide` — the slide's description fields (see slide schema below)\n"
        f"- `assets` — a dict keyed by asset id (e.g. `\"asset_abc12345\"`); "
        f"each value is that asset's description fields (see asset schema "
        f"below). Include exactly the asset ids listed for that bundle "
        f"above. If a bundle has 0 assets, output `\"assets\": {{}}`.\n\n"
        f"```json\n{sample_block}```\n\n"
        f"Return EXACTLY {n} top-level entries. Use the asset ids as listed "
        f"above — do NOT rename them. Output ONLY the JSON object. No "
        f"commentary, no markdown code fences, no prose before or after.\n\n"
        f"---\n\n"
        f"## Slide description schema (use under each `slide` key)\n\n"
        f"{slide_prompt}\n\n"
        f"---\n\n"
        f"## Asset description schema (use under each `assets.<id>` key)\n\n"
        f"{asset_prompt}\n"
    )


def _bulk_instructions(kind: str, n: int, per_item_prompt: str) -> str:
    item_name = "image" if kind == "asset" else "slide preview"
    if kind == "asset":
        sample_block = (
            '{\n'
            '  "01": {\n'
            '    "kind": "photo",\n'
            '    "subject": "...",\n'
            '    "depicts": "...",\n'
            '    "feel": "warm",\n'
            '    "composition": "centered",\n'
            '    "colors": ["navy", "white"],\n'
            '    "scope": ["generic"],\n'
            '    "suitable_for": ["team"],\n'
            '    "notes": "",\n'
            '    "interpretation": ""\n'
            '  },\n'
            '  "02": { "kind": "photo", "...": "same fields" }\n'
            '}\n'
        )
    else:
        sample_block = (
            '{\n'
            '  "01": {\n'
            '    "intent": "...",\n'
            '    "feel": "formal",\n'
            '    "suitable_for": ["opener"],\n'
            '    "notes": "",\n'
            '    "interpretation": ""\n'
            '  },\n'
            '  "02": { "intent": "...", "...": "same fields" }\n'
            '}\n'
        )
    return (
        f"# Bulk describe batch — {n} {item_name}s\n\n"
        f"You will see {n} {item_name}s numbered 01 through {n:02d}. For each, "
        f"produce a description following the schema in the second half of "
        f"this file.\n\n"
        f"## Output format\n\n"
        f"Return ONE JSON object. Top-level keys are the quoted 2-digit ids "
        f"(`\"01\"`, `\"02\"`, ..., `\"{n:02d}\"`); each value is an object "
        f"holding that item's fields.\n\n"
        f"```json\n{sample_block}```\n\n"
        f"Return EXACTLY {n} entries, one per item. Do NOT skip any. Output "
        f"ONLY the JSON object. No commentary, no markdown code fences, no "
        f"prose before or after.\n\n"
        f"---\n\n"
        f"## Per-item description schema\n\n"
        f"{per_item_prompt}\n"
    )


# Cap on how many describe batches we keep on disk. Each batch holds a
# zip (slide PNGs + asset binaries) plus a manifest; at ~2 MB/bundle
# they accumulate quickly in long-running workspaces. The /api/batches
# list endpoint still caps at 20 for safety, but in practice we prune
# eagerly here. Losing a batch before applying it just means
# regenerating — cheap.
_KEEP_BATCHES = 5


def _prune_old_batches(keep: int = _KEEP_BATCHES) -> int:
    """Delete all but the `keep` most-recent batch dirs. Returns count removed.

    Batch ids are `%Y%m%d-%H%M%S` so lexicographic sort == chronological.
    """
    if not BATCHES_DIR.exists():
        return 0
    dirs = sorted(
        [d for d in BATCHES_DIR.iterdir() if d.is_dir()],
        key=lambda d: d.name,
        reverse=True,
    )
    removed = 0
    for stale in dirs[keep:]:
        try:
            shutil.rmtree(stale)
            removed += 1
        except OSError:
            pass
    return removed


@app.post("/api/batch/create")
def api_batch_create():
    body = request.get_json(force=True) or {}
    kind = body.get("kind", "asset")
    try:
        count = max(1, min(20, int(body.get("count", 10))))
    except (TypeError, ValueError):
        return jsonify({"error": "count must be an integer"}), 400
    if kind not in ("asset", "slide", "slide_with_assets"):
        return jsonify({
            "error": "kind must be 'asset', 'slide', or 'slide_with_assets'",
        }), 400

    if kind == "asset":
        candidates = list(cli_mod.iter_asset_yamls())
    else:
        # Both 'slide' and 'slide_with_assets' start from pending slides.
        candidates = list(cli_mod.iter_slide_yamls())
    pending = [p for p in candidates if _safe_read(p).get("status", "pending") == "pending"]
    selected = pending[:count]
    if not selected:
        return jsonify({"error": f"no pending {kind}s"}), 400

    BATCHES_DIR.mkdir(exist_ok=True)
    batch_id = datetime.now().strftime("%Y%m%d-%H%M%S")
    batch_dir = BATCHES_DIR / batch_id
    batch_dir.mkdir(exist_ok=True)

    manifest: dict = {"kind": kind, "created": batch_id, "items": {}}
    skipped: list[dict] = []
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        added = 0
        for ypath in selected:
            rel = str(ypath.relative_to(HERE))
            try:
                if kind == "asset":
                    binary = _asset_binary(ypath)
                    if binary is None:
                        skipped.append({"yaml": rel, "reason": "asset binary missing"})
                        continue
                    if not binary.exists() or binary.stat().st_size == 0:
                        skipped.append({"yaml": rel, "reason": f"asset binary unreadable: {binary.name}"})
                        continue
                    ext = binary.suffix.lstrip(".") or "png"
                    key = f"{added + 1:02d}"
                    zf.write(binary, f"{key}.{ext}")
                    manifest["items"][key] = rel
                    added += 1
                elif kind == "slide":
                    slide_pptx = ypath.with_suffix(".pptx")
                    if not slide_pptx.exists():
                        skipped.append({"yaml": rel, "reason": "slide fragment .pptx missing"})
                        continue
                    png = _ensure_slide_png(slide_pptx)
                    if png is None:
                        available = cli_mod.available_renderers()
                        if not available:
                            reason = (
                                "no slide renderer available — install "
                                "LibreOffice, or use PowerPoint on Windows"
                            )
                        else:
                            reason = (
                                f"slide rendering failed (tried: "
                                f"{', '.join(available)})"
                            )
                        skipped.append({"yaml": rel, "reason": reason})
                        continue
                    key = f"{added + 1:02d}"
                    zf.write(png, f"{key}.png")
                    manifest["items"][key] = rel
                    added += 1
                else:
                    # slide_with_assets: one numbered folder per slide,
                    # containing slide.png + each asset binary. Manifest
                    # value is a nested dict so apply can route both
                    # halves of the LLM response.
                    slide_pptx = ypath.with_suffix(".pptx")
                    if not slide_pptx.exists():
                        skipped.append({"yaml": rel, "reason": "slide fragment .pptx missing"})
                        continue
                    png = _ensure_slide_png(slide_pptx)
                    if png is None:
                        available = cli_mod.available_renderers()
                        reason = (
                            "no slide renderer available — install "
                            "LibreOffice, or use PowerPoint on Windows"
                            if not available
                            else f"slide rendering failed (tried: {', '.join(available)})"
                        )
                        skipped.append({"yaml": rel, "reason": reason})
                        continue
                    key = f"{added + 1:02d}"
                    zf.write(png, f"{key}/slide.png")
                    asset_yamls = _assets_for_slide(ypath)
                    assets_map: dict[str, str] = {}
                    for ap in asset_yamls:
                        a_data = _safe_read(ap)
                        a_id = a_data.get("id", "")
                        if not a_id:
                            continue
                        a_bin = _asset_binary(ap)
                        if a_bin is None or not a_bin.exists() or a_bin.stat().st_size == 0:
                            # Skip this individual asset but keep the bundle.
                            continue
                        ext = a_bin.suffix.lstrip(".") or "bin"
                        zf.write(a_bin, f"{key}/{a_id}.{ext}")
                        assets_map[a_id] = str(ap.relative_to(HERE))
                    manifest["items"][key] = {"slide": rel, "assets": assets_map}
                    added += 1
            except OSError as e:
                skipped.append({"yaml": rel, "reason": f"OS error: {e}"})
            except Exception as e:
                skipped.append({"yaml": rel, "reason": f"unexpected: {type(e).__name__}: {e}"})

        if kind == "slide_with_assets":
            slide_prompt = (HERE / "prompts" / "describe_slide.md").read_text(encoding="utf-8")
            asset_prompt = (HERE / "prompts" / "describe_asset.md").read_text(encoding="utf-8")
            zf.writestr(
                "instructions.md",
                _bulk_instructions_slide_with_assets(
                    added, slide_prompt, asset_prompt, manifest["items"],
                ),
            )
        else:
            per_item_prompt = (HERE / "prompts" / f"describe_{kind}.md").read_text(
                encoding="utf-8"
            )
            zf.writestr(
                "instructions.md",
                _bulk_instructions(kind, added, per_item_prompt),
            )

    if not manifest["items"]:
        shutil.rmtree(batch_dir, ignore_errors=True)
        return jsonify({
            "error": f"no {kind} previews could be generated",
            "skipped": skipped,
        }), 500

    (batch_dir / "batch.zip").write_bytes(zip_buf.getvalue())
    (batch_dir / "manifest.json").write_text(
        json_mod.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    pruned = _prune_old_batches()

    return jsonify({
        "batch_id": batch_id,
        "kind": kind,
        "count": added,
        "requested": count,
        "items": manifest["items"],
        "skipped": skipped,
        "download_url": f"/api/batch/{batch_id}/download",
        "pruned": pruned,
    })


@app.get("/api/batch/<batch_id>/download")
def api_batch_download(batch_id):
    batch_dir = BATCHES_DIR / batch_id
    zip_path = batch_dir / "batch.zip"
    if not zip_path.exists():
        abort(404, "batch not found")
    return send_file(
        zip_path,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"batch_{batch_id}.zip",
    )


def _normalize_key_to_int(k) -> int | None:
    """Extract any digits from a key (str/int/bytes) and return as int."""
    if isinstance(k, int):
        return k
    s = str(k)
    digits = "".join(c for c in s if c.isdigit())
    return int(digits) if digits else None


def _find_items_dict(parsed):
    """Drill into common wrapper shapes the LLM might emit.

    - dict of dicts → use directly
    - list of dicts → key by 1-based position
    - {wrapper: dict|list} → descend
    """
    if isinstance(parsed, list):
        return {f"{i:02d}": v for i, v in enumerate(parsed, 1)}
    if not isinstance(parsed, dict):
        return None
    if parsed and all(isinstance(v, dict) for v in parsed.values()):
        return parsed
    if len(parsed) == 1:
        inner = next(iter(parsed.values()))
        if isinstance(inner, (dict, list)):
            return _find_items_dict(inner)
    return parsed


@app.post("/api/batch/<batch_id>/apply")
def api_batch_apply(batch_id):
    batch_dir = BATCHES_DIR / batch_id
    manifest_path = batch_dir / "manifest.json"
    if not manifest_path.exists():
        return jsonify({"error": "batch not found"}), 404
    manifest = json_mod.loads(manifest_path.read_text(encoding="utf-8"))

    body = request.get_json(force=True) or {}
    txt = _strip_yaml_fence(body.get("text", ""))
    try:
        parsed = json_mod.loads(txt) if txt else None
    except Exception as e:
        return jsonify({"error": f"JSON parse error: {e}"}), 400
    if not isinstance(parsed, (dict, list)):
        return jsonify({"error": "expected a JSON object or array at top level"}), 400

    items_dict = _find_items_dict(parsed) or {}
    found_keys = [str(k) for k in items_dict.keys()]
    by_int: dict[int, dict] = {}
    for k, v in items_dict.items():
        n = _normalize_key_to_int(k)
        if n is not None and isinstance(v, dict):
            by_int[n] = v

    batch_kind = manifest.get("kind", "asset")
    results = []
    for key, manifest_entry in manifest["items"].items():
        n = _normalize_key_to_int(key)
        entry = by_int.get(n) if n is not None else None

        # slide_with_assets: manifest_entry is a dict {slide, assets},
        # and entry is expected to mirror that nested shape.
        if batch_kind == "slide_with_assets" and isinstance(manifest_entry, dict):
            slide_rel = manifest_entry.get("slide", "")
            assets_map = manifest_entry.get("assets") or {}
            if entry is None:
                results.append({
                    "id": key, "yaml": slide_rel, "status": "no-match",
                    "errors": ["LLM response missing this bundle key"],
                })
                continue
            slide_fields = entry.get("slide") if isinstance(entry.get("slide"), dict) else None
            asset_fields = entry.get("assets") if isinstance(entry.get("assets"), dict) else {}

            # Apply slide half.
            if slide_fields is None:
                results.append({
                    "id": key, "yaml": slide_rel, "status": "no-match",
                    "errors": ["bundle entry missing 'slide' object"],
                })
            else:
                try:
                    p = _yaml_for_rel(slide_rel)
                    errs, status = _save_and_validate(p, slide_fields)
                    results.append({
                        "id": key, "yaml": slide_rel, "status": status, "errors": errs,
                    })
                except Exception as e:
                    results.append({
                        "id": key, "yaml": slide_rel, "status": "error",
                        "errors": [str(e)],
                    })

            # Apply each asset half. Skip with a notice if already done —
            # avoids overwriting a description the user already validated
            # via the regular asset bulk mode or by hand.
            for asset_id, asset_rel in assets_map.items():
                sub_key = f"{key}/{asset_id}"
                af = asset_fields.get(asset_id) if isinstance(asset_fields, dict) else None
                if not isinstance(af, dict):
                    results.append({
                        "id": sub_key, "yaml": asset_rel, "status": "no-match",
                        "errors": [f"bundle entry missing assets.{asset_id}"],
                    })
                    continue
                try:
                    ap = _yaml_for_rel(asset_rel)
                    existing = _safe_read(ap)
                    if existing.get("status") == "done":
                        results.append({
                            "id": sub_key, "yaml": asset_rel, "status": "skipped",
                            "errors": [
                                f"asset {asset_id} already described as 'done'; "
                                f"keeping existing description, skipping bundle override"
                            ],
                        })
                        continue
                    errs, status = _save_and_validate(ap, af)
                    results.append({
                        "id": sub_key, "yaml": asset_rel, "status": status, "errors": errs,
                    })
                except Exception as e:
                    results.append({
                        "id": sub_key, "yaml": asset_rel, "status": "error",
                        "errors": [str(e)],
                    })
            continue

        # Flat batch kinds (asset, slide): manifest_entry is a path string.
        rel = manifest_entry if isinstance(manifest_entry, str) else str(manifest_entry)
        if entry is None:
            results.append({
                "id": key, "yaml": rel, "status": "no-match",
                "errors": ["LLM response missing this key"],
            })
            continue
        try:
            p = _yaml_for_rel(rel)
        except Exception as e:
            results.append({
                "id": key, "yaml": rel, "status": "error", "errors": [str(e)],
            })
            continue
        errs, status = _save_and_validate(p, entry)
        results.append({
            "id": key, "yaml": rel, "status": status, "errors": errs,
        })
    return jsonify({
        "batch_id": batch_id,
        "results": results,
        "found_keys": found_keys,
        "matched": sum(1 for r in results if r["status"] not in ("no-match",)),
    })


@app.get("/api/batch/<batch_id>")
def api_batch_info(batch_id):
    mf = BATCHES_DIR / batch_id / "manifest.json"
    if not mf.exists():
        abort(404, "batch not found")
    return jsonify(json_mod.loads(mf.read_text(encoding="utf-8")))


@app.get("/api/batches")
def api_batches_list():
    if not BATCHES_DIR.exists():
        return jsonify({"batches": []})
    out = []
    for d in sorted(BATCHES_DIR.iterdir(), reverse=True):
        if not d.is_dir():
            continue
        mf = d / "manifest.json"
        if not mf.exists():
            continue
        m = json_mod.loads(mf.read_text(encoding="utf-8"))
        out.append({
            "batch_id": d.name,
            "kind": m.get("kind"),
            "count": len(m.get("items", {})),
        })
    # Cap matches _KEEP_BATCHES so the dropdown can't surface a batch
    # that's already been pruned from disk. Defense-in-depth — disk
    # pruning runs on every create, so out should already be ≤_KEEP_BATCHES.
    return jsonify({"batches": out[:_KEEP_BATCHES]})


@app.get("/preview")
def preview():
    rel = request.args.get("yaml", "")
    p = _yaml_for_rel(rel)
    if _kind_of(p) == "slide":
        slide_pptx = p.with_suffix(".pptx")
        if not slide_pptx.exists():
            abort(404, "slide pptx missing")
        png = _ensure_slide_png(slide_pptx)
        if png is None:
            abort(503, "qlmanage not available — install or run on macOS")
        return send_file(png, mimetype="image/png")
    binary = _asset_binary(p)
    if binary is None:
        abort(404, "asset binary missing")
    return send_file(binary)


# ---------------------------------------------------------------------------
# Compose flow — filter KB, build prompt bundle, run compose from a plan
# ---------------------------------------------------------------------------


FILTER_DIMENSIONS = {
    "templates": ("feel", "suitable_for"),
    "assets": ("kind", "feel", "composition", "suitable_for", "scope", "colors"),
}

BRAND_PATH = HERE / "brand.md"
PRESETS_DIR = WORKSPACE / "_presets"
_PRESET_NAME_RE = __import__("re").compile(r"^[A-Za-z0-9][A-Za-z0-9_\- ]{0,60}$")


def _read_brand() -> str:
    if BRAND_PATH.exists():
        return BRAND_PATH.read_text(encoding="utf-8")
    return ""


def _write_brand(text: str) -> None:
    BRAND_PATH.write_text(text, encoding="utf-8")


def _preset_path(name: str) -> Path:
    if not _PRESET_NAME_RE.match(name):
        abort(400, "invalid preset name")
    return PRESETS_DIR / f"{name}.md"


def _list_presets() -> list[dict]:
    if not PRESETS_DIR.exists():
        return []
    out: list[dict] = []
    for p in sorted(PRESETS_DIR.glob("*.md")):
        out.append({
            "name": p.stem,
            "preview": p.read_text(encoding="utf-8").strip().splitlines()[0:1] or [""],
        })
    # Flatten preview lists to a single-line string.
    for item in out:
        item["preview"] = item["preview"][0] if item["preview"] else ""
    return out


def _collect_descriptions() -> tuple[list[dict], list[dict]]:
    slides: list[dict] = []
    for p in cli_mod.iter_slide_yamls():
        d = _safe_read(p)
        if not d:
            continue
        d["_yaml_path"] = p
        slides.append(d)
    assets: list[dict] = []
    for p in cli_mod.iter_asset_yamls():
        d = _safe_read(p)
        if not d:
            continue
        d["_yaml_path"] = p
        assets.append(d)
    return slides, assets


def _collect_filter_options() -> dict:
    slides, assets = _collect_descriptions()

    def collect(items: list[dict], fields: tuple[str, ...]) -> dict:
        out: dict[str, set[str]] = {f: set() for f in fields}
        for it in items:
            for f in fields:
                v = it.get(f)
                if v is None:
                    continue
                if isinstance(v, list):
                    for x in v:
                        if isinstance(x, str) and x:
                            out[f].add(x)
                elif isinstance(v, str) and v:
                    out[f].add(v)
        return {k: sorted(v) for k, v in out.items()}

    return {
        "templates": {
            "options": collect(slides, FILTER_DIMENSIONS["templates"]),
            "total": len(slides),
        },
        "assets": {
            "options": collect(assets, FILTER_DIMENSIONS["assets"]),
            "total": len(assets),
        },
    }


def _matches_filters(item: dict, filters: dict) -> bool:
    for field, allowed in filters.items():
        if not allowed:
            continue
        v = item.get(field)
        if v is None:
            return False
        if isinstance(v, list):
            if not any(x in allowed for x in v):
                return False
        elif v not in allowed:
            return False
    return True


def _filter_kb(filters: dict) -> tuple[list[dict], list[dict]]:
    slides, assets = _collect_descriptions()
    tpl_filters = filters.get("templates") or {}
    ast_filters = filters.get("assets") or {}
    slides_out = [s for s in slides if _matches_filters(s, tpl_filters)]
    assets_out = [a for a in assets if _matches_filters(a, ast_filters)]
    return slides_out, assets_out


def _read_skill_md() -> str:
    return (cli_mod.CONSUMER / "SKILL.md").read_text(encoding="utf-8")


def _format_user_assets_section(user_meta: dict) -> str:
    """Brief.md sub-section listing user-supplied assets and how to treat
    them. Returns the empty string when none are attached.
    """
    if not user_meta:
        return ""
    lines = [
        "## User-supplied assets — IMPORTANT",
        "",
        "The user attached the following assets to THIS request. They "
        "are the primary intent and SHOULD be used in the deck wherever "
        "they fit. Each file lives at `user_assets/<id>.<ext>` in this "
        "bundle as a LOW-RES preview; the user's machine holds the "
        "original at full resolution and will splice the original in at "
        "compose time. The dimensions below are the originals'.",
        "",
        "Reference them in the plan exactly the same way you reference "
        "any catalog asset — `\"<slot>\": \"<id>\"` for image slots, or "
        "`{\"atom\": \"<id>\", ...}` inside a compose-mode shape. The id "
        "format matches catalog assets on purpose so the compose "
        "pipeline resolves them transparently.",
        "",
    ]
    for aid, entry in sorted(user_meta.items()):
        kind = entry.get("kind", "?")
        ext = entry.get("ext", "?")
        fname = entry.get("filename", "?")
        size_kb = max(1, (entry.get("size_bytes") or 0) // 1024)
        dims = ""
        w, h = entry.get("width"), entry.get("height")
        if w and h:
            dims = f" {w}x{h}px"
        lines.append(
            f"- `{aid}` — {kind} ({ext}){dims}, "
            f"{size_kb} KB original — original filename: `{fname}`"
        )
    lines += [
        "",
        "### How to treat these",
        "",
        "- The user expects these assets to appear in the output. Treat "
        "them as a stronger signal than KB catalog matches.",
        "- If a user asset fits a slot, use it — even if a KB asset "
        "would score better on `feel` / `colors`. The user's intent "
        "trumps the descriptive index here.",
        "- If a user asset CANNOT fit any slot in your plan (wrong "
        "aspect, no semantic match, etc.), prefer in this order:",
        "  1. Pick a different template whose layout exposes a slot "
        "this asset fits.",
        "  2. Use compose-mode (free-form atoms) to place the user "
        "asset alongside other shapes you control.",
        "  3. Fall back to a similar KB catalog asset and explicitly "
        "explain in the brief response that the user asset didn't fit.",
        "  4. Last resort: omit the user asset, but flag it.",
        "- The previews are LOW-RES (max 800px long side, possibly "
        "re-encoded). Judge subject / composition / colors from them, "
        "but don't reason about pixel-level detail.",
        "- These assets do NOT carry descriptions. You see the file "
        "itself plus the brief — combine the two to infer intent.",
        "",
    ]
    return "\n".join(lines)


def _format_brief(brief: str, user_meta: dict | None = None) -> str:
    user_section = _format_user_assets_section(user_meta or {})
    return (
        "# Deck brief\n\n"
        f"{brief.strip() or '(no brief supplied)'}\n\n"
        "---\n\n"
        + (user_section + "---\n\n" if user_section else "")
        + "# Output rules\n\n"
        "- Read SKILL.md for the three-command surface and plan format.\n"
        "- Pick templates and assets ONLY from the attached `index.json`.\n"
        "- Return ONLY a JSON array. No prose, no markdown fences.\n"
        "- Each entry: {\"template\": \"<id>\", \"slots\": { ... }}.\n"
        "- Slot keys must match the template's declared slot ids.\n"
        "- Image slot values: an asset id from index.json (not a file path).\n"
        "- Text slot values: plain string, respect each slot's `max_chars`.\n"
        "- Bullets slot values: array of strings.\n"
        "- If no asset fits a slot, omit the slot rather than forcing one.\n"
        "\n"
        "## Helpers in this bundle\n"
        "\n"
        "Read-only Python utilities under `helpers/` — invoke them from the\n"
        "bundle root (stdlib + pyyaml; exit codes 0 ok / 1 empty-or-fail / 2\n"
        "bad input). They expose data; you decide the picks.\n"
        "\n"
        "- `python helpers/kb_summary.py` — facet counts; start here on a "
        "fresh bundle\n"
        "- `python helpers/kb_filter.py templates --feel formal "
        "--suitable_for opener`\n"
        "- `python helpers/kb_filter.py assets --kind photo --text \"team\"`\n"
        "- `python helpers/kb_inspect.py <id>` — denormalized view; inlines "
        "inventory atoms with their descriptions\n"
        "- `python helpers/kb_lint.py < plan.json` — pre-flight validator "
        "(slot ids, max_chars, bullet glyphs, missing assets, degraded "
        "shapes)\n"
        "- `python helpers/kb_themes.py [--resolve-role <role> --for-deck "
        "<deck>]` — per-deck theme + alias-aware role resolution\n"
        "- `python helpers/kb_budget.py <template_id> <slot_id> \"text\"` — "
        "one-shot text-budget check\n"
        "\n"
        "See `helpers/README.md` for full docs.\n"
        "\n"
        "## v4 capabilities — what's new\n"
        "\n"
        "- Each template carries `theme_colors` and `fonts` showing the\n"
        "  source deck's actual palette + theme fonts. Pick templates\n"
        "  whose theme is close to brand policy when possible.\n"
        "- Each template has `inventory` listing structured atoms it\n"
        "  carries (tables, callouts, charts, smartart). Picking the\n"
        "  template brings those atoms along for free.\n"
        "- Asset kinds now include `vector`, `table`, `chart`, `callout`,\n"
        "  `freeform`, `smartart` — filter on them via `kind=table` etc.\n"
        "- Each asset has `colors_hex` (actual hex from binary inspection)\n"
        "  alongside the human-readable `colors` words.\n"
        "\n"
        "## Slot value polymorphism — accepted but partially honored\n"
        "\n"
        "Slot values can be more than plain strings/arrays/asset ids:\n"
        "  - {\"text\": \"…\", \"color_role\": \"accent\", \"bold\": true}\n"
        "  - {\"runs\": [{\"text\": \"X\", \"bold\": true}, {\"text\": \" Y\"}]}\n"
        "  - {\"asset\": \"asset_<id>\", \"recolor\": {\"#ff0000\": \"accent\"}}\n"
        "\n"
        "Current build: styling fields are accepted without crashing but\n"
        "DROPPED at compose time with a one-line warning. Phase D will\n"
        "honor them. Until then, prefer plain strings/arrays/asset ids\n"
        "unless you have a specific reason to flag styling intent.\n"
        "\n"
        "## Compose-mode entries — accepted but skipped\n"
        "\n"
        "Plan entries of shape {\"compose\": true, \"layout\": \"…\",\n"
        "\"shapes\": [...]} let you assemble a slide from atoms. The\n"
        "engine currently SKIPS these entries with a warning (full\n"
        "support lands in Phase D). For a slide to render today it\n"
        "must use {\"template\": \"…\", \"slots\": {...}} shape.\n"
        "\n"
        "## Bullets — DO NOT prepend bullet glyphs\n"
        "\n"
        "PowerPoint templates apply bullets via layout formatting. If you\n"
        "prepend a literal `•`, `-`, or `*` to lines, the rendered slide\n"
        "shows two bullets per line (e.g. `•• My point`).\n"
        "\n"
        "Correct ways to produce a bulleted list:\n"
        "\n"
        "- For a slot of `kind: bullets` — pass an array of plain strings.\n"
        "  GOOD: `\"body\": [\"First point\", \"Second point\"]`\n"
        "  BAD:  `\"body\": [\"• First point\", \"• Second point\"]`\n"
        "- For a `kind: text` slot that visually behaves as bullets in the\n"
        "  template — pass plain strings joined by `\\n`, NO leading glyph.\n"
        "  GOOD: `\"subtitle\": \"First point\\nSecond point\"`\n"
        "  BAD:  `\"subtitle\": \"• First point\\n• Second point\"`\n"
    )


def _build_prompt_bundle_zip(slides: list[dict], assets: list[dict], brief: str) -> bytes:
    clean_slides = [{k: v for k, v in s.items() if not k.startswith("_")} for s in slides]
    clean_assets = [{k: v for k, v in a.items() if not k.startswith("_")} for a in assets]
    index = cli_mod.build_index(clean_slides, clean_assets)
    user_meta = _read_user_meta(USER_STAGED_DIR)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("SKILL.md", _read_skill_md())
        brand = _read_brand().strip()
        if brand:
            zf.writestr("brand.md", brand + "\n")
        zf.writestr("index.json", json_mod.dumps(index, indent=2, ensure_ascii=False))
        zf.writestr("brief.md", _format_brief(brief, user_meta))
        # v4: per-deck theme.yaml. Filtered KB may have culled some
        # decks entirely — only ship themes whose deck still has at
        # least one template in the bundle (so the agent's bundle
        # mirrors what they actually see).
        decks_in_bundle = {
            (s.get("sources") or [{}])[0].get("deck", "")
            for s in clean_slides
        }
        decks_in_bundle.discard("")
        for theme_yaml in sorted((cli_mod.WORKSPACE / "decks").glob("*/theme.yaml")):
            if theme_yaml.parent.name not in decks_in_bundle:
                continue
            zf.writestr(
                f"decks/{theme_yaml.parent.name}/theme.yaml",
                theme_yaml.read_text(encoding="utf-8"),
            )
        # Agent-side helpers: read-only kb_* scripts the agent invokes
        # against the bundle from its own working dir (filter the catalog,
        # inspect entries, lint a draft plan). See helpers/README.md inside
        # the bundle. Stdlib + pyyaml only.
        helpers_dir = cli_mod.CONSUMER / "helpers"
        if helpers_dir.exists():
            for hp in sorted(helpers_dir.iterdir()):
                if hp.is_file() and not hp.name.startswith("."):
                    zf.write(hp, f"helpers/{hp.name}")
        # User-supplied assets: low-res previews + manifest with original
        # dimensions. Full-res originals stay on the user's machine and
        # are spliced into compose-run staging when the plan comes back.
        if user_meta:
            manifest_for_zip: dict = {}
            for aid, entry in user_meta.items():
                ext = entry.get("ext") or "bin"
                src = USER_STAGED_DIR / f"{aid}.{ext}"
                if not src.exists():
                    continue
                # Write a low-res copy directly into the zip via a tmp
                # file so we don't materialize huge images in RAM.
                with tempfile.NamedTemporaryFile(
                    delete=False, suffix=f".{ext}",
                ) as tf:
                    low = Path(tf.name)
                try:
                    _make_low_res(src, low, entry.get("kind", "image"))
                    zf.write(low, f"user_assets/{aid}.{ext}")
                finally:
                    low.unlink(missing_ok=True)
                manifest_for_zip[aid] = {
                    "id": aid,
                    "filename": entry.get("filename"),
                    "kind": entry.get("kind"),
                    "ext": ext,
                    "size_bytes": entry.get("size_bytes"),
                    "width": entry.get("width"),
                    "height": entry.get("height"),
                }
            zf.writestr(
                "user_assets/manifest.json",
                json_mod.dumps(
                    {"note": (
                        "User-supplied assets attached to this request. "
                        "The files here are LOW-RES previews; the "
                        "user's machine holds the full-resolution "
                        "originals and will splice them in at compose "
                        "time. See brief.md for usage guidance."
                    ),
                     "assets": manifest_for_zip},
                    indent=2, ensure_ascii=False,
                ),
            )
    blob = buf.getvalue()
    # Move staged → bundle AFTER the zip is finalized: the bundle/ dir
    # is now the canonical snapshot compose-run will use to resolve any
    # user asset ids the agent references in the plan.
    if user_meta:
        _move_staged_to_bundle()
    return blob


def _flat_prompt_text(slides: list[dict], assets: list[dict], brief: str) -> str:
    clean_slides = [{k: v for k, v in s.items() if not k.startswith("_")} for s in slides]
    clean_assets = [{k: v for k, v in a.items() if not k.startswith("_")} for a in assets]
    index = cli_mod.build_index(clean_slides, clean_assets)
    sections: list[str] = []
    brand = _read_brand().strip()
    if brand:
        sections.append("=== brand.md ===\n" + brand)
    sections.append("=== SKILL.md ===\n" + _read_skill_md())
    sections.append("=== index.json ===\n" + json_mod.dumps(index, indent=2, ensure_ascii=False))
    user_meta = _read_user_meta(USER_STAGED_DIR)
    sections.append("=== brief.md ===\n" + _format_brief(brief, user_meta))
    if user_meta:
        sections.append(
            "=== user_assets note ===\n"
            "User-supplied asset BINARIES are not included in this "
            "flat-text view. Use the .zip bundle path to see them; "
            "metadata is listed in brief.md above."
        )
    return "\n\n".join(sections) + "\n"


def _stage_compose_bundle(staging: Path) -> None:
    """Stage a minimal reader.py-compatible bundle (full KB) under `staging`.

    Produces the same on-disk layout `cli build` does (sans SKILL.md /
    brand.md, which compose-run doesn't need): reader.py, index.json,
    templates/<id>/slide.pptx + meta.yaml, assets/<id>.<ext> + .yaml,
    decks/<deck>/theme.yaml (so D5 cross-deck remap can resolve themes).
    """
    slides, assets = _collect_descriptions()
    (staging / "reader.py").write_text(
        (cli_mod.CONSUMER / "reader.py").read_text(encoding="utf-8"),
        encoding="utf-8",
    )
    tpl_dir = staging / "templates"
    for sd in slides:
        tid = sd["id"]
        yaml_path: Path = sd["_yaml_path"]
        slide_pptx = yaml_path.with_suffix(".pptx")
        if not slide_pptx.exists():
            continue
        d = tpl_dir / tid
        d.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(slide_pptx, d / "slide.pptx")
        clean = {k: v for k, v in sd.items() if not k.startswith("_")}
        (d / "meta.yaml").write_text(
            yaml.safe_dump(clean, sort_keys=False, allow_unicode=True),
            encoding="utf-8",
        )
    ast_dir = staging / "assets"
    ast_dir.mkdir(parents=True, exist_ok=True)
    for ad in assets:
        aid = ad["id"]
        yaml_path: Path = ad["_yaml_path"]
        binary = _asset_binary(yaml_path)
        if binary is None:
            continue
        ext = binary.suffix.lstrip(".") or "bin"
        shutil.copyfile(binary, ast_dir / f"{aid}.{ext}")
        clean = {k: v for k, v in ad.items() if not k.startswith("_")}
        (ast_dir / f"{aid}.yaml").write_text(
            yaml.safe_dump(clean, sort_keys=False, allow_unicode=True),
            encoding="utf-8",
        )

    # index.json — required by reader.py's load_index(). Same payload
    # build_index() produces during `cli build`, minus the _yaml_path
    # sidecar that's only meaningful inside the workspace.
    clean_slides = [{k: v for k, v in s.items() if not k.startswith("_")} for s in slides]
    clean_assets = [{k: v for k, v in a.items() if not k.startswith("_")} for a in assets]
    index = cli_mod.build_index(clean_slides, clean_assets)
    (staging / "index.json").write_text(
        json_mod.dumps(index, indent=2, ensure_ascii=False), encoding="utf-8",
    )

    # Per-deck theme.yaml — needed for D5 cross-deck colour remap +
    # v4.1 font remap when an atom is pulled from a foreign deck onto
    # a host. Only ship themes whose deck still has at least one
    # described template (mirrors _build_prompt_bundle_zip's logic).
    decks_in_bundle = {
        (s.get("sources") or [{}])[0].get("deck", "") for s in clean_slides
    }
    decks_in_bundle.discard("")
    decks_dir = staging / "decks"
    for theme_yaml in sorted((cli_mod.WORKSPACE / "decks").glob("*/theme.yaml")):
        if theme_yaml.parent.name not in decks_in_bundle:
            continue
        deck_out = decks_dir / theme_yaml.parent.name
        deck_out.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(theme_yaml, deck_out / "theme.yaml")


@app.get("/api/compose/options")
def api_compose_options():
    return jsonify(_collect_filter_options())


@app.post("/api/compose/preview")
def api_compose_preview():
    body = request.get_json(force=True) or {}
    filters = body.get("filters") or {}
    slides, assets = _filter_kb(filters)
    return jsonify({
        "templates": len(slides),
        "assets": len(assets),
        "template_ids": [s["id"] for s in slides],
        "asset_ids": [a["id"] for a in assets],
    })


@app.post("/api/compose/bundle")
def api_compose_bundle():
    body = request.get_json(force=True) or {}
    filters = body.get("filters") or {}
    brief = body.get("brief") or ""
    slides, assets = _filter_kb(filters)
    if not slides and not assets:
        return jsonify({"error": "filters match nothing — broaden them"}), 400
    blob = _build_prompt_bundle_zip(slides, assets, brief)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    return send_file(
        io.BytesIO(blob),
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"prompt-bundle-{ts}.zip",
    )


@app.post("/api/compose/text")
def api_compose_text():
    body = request.get_json(force=True) or {}
    filters = body.get("filters") or {}
    brief = body.get("brief") or ""
    slides, assets = _filter_kb(filters)
    return jsonify({"text": _flat_prompt_text(slides, assets, brief)})


@app.get("/api/vocab")
def api_vocab():
    return jsonify(cli_mod.VOCAB)


@app.get("/api/compose/brand")
def api_compose_brand_get():
    return jsonify({"text": _read_brand(), "path": str(BRAND_PATH.relative_to(HERE))})


@app.put("/api/compose/brand")
def api_compose_brand_put():
    body = request.get_json(force=True) or {}
    text = body.get("text", "")
    if not isinstance(text, str):
        return jsonify({"error": "text must be a string"}), 400
    _write_brand(text)
    return jsonify({"ok": True, "chars": len(text)})


@app.get("/api/compose/presets")
def api_compose_presets_list():
    return jsonify({"presets": _list_presets()})


@app.get("/api/compose/preset")
def api_compose_preset_get():
    name = request.args.get("name", "")
    p = _preset_path(name)
    if not p.exists():
        abort(404, "preset not found")
    return jsonify({"name": name, "text": p.read_text(encoding="utf-8")})


@app.put("/api/compose/preset")
def api_compose_preset_put():
    body = request.get_json(force=True) or {}
    name = body.get("name", "")
    text = body.get("text", "")
    if not isinstance(text, str):
        return jsonify({"error": "text must be a string"}), 400
    p = _preset_path(name)
    PRESETS_DIR.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding="utf-8")
    return jsonify({"ok": True, "name": name})


@app.delete("/api/compose/preset")
def api_compose_preset_delete():
    name = request.args.get("name", "")
    p = _preset_path(name)
    if p.exists():
        p.unlink()
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# User-supplied assets — per-request attachments (separate from KB)
# ---------------------------------------------------------------------------


# Allow common raster images, SVG vectors, and XML atom fragments. The
# user said they want tables/charts/etc. to work too — those are all
# XML fragments in this codebase, so a single .xml acceptor covers them.
_USER_EXT_KIND = {
    ".png":  "image",  ".jpg":  "image",  ".jpeg": "image",
    ".webp": "image",  ".gif":  "image",
    ".svg":  "vector",
    ".xml":  "atom",
}
_USER_MAX_FILE_BYTES = 20 * 1024 * 1024   # 20 MB / file
_USER_MAX_TOTAL_BYTES = 100 * 1024 * 1024  # 100 MB across all staged
_USER_MAX_FILES = 30
_USER_LOW_RES_LONG_SIDE = 800  # px — for the low-res copy shipped in the zip


def _clear_user_staged_on_startup() -> None:
    """Per user preference: staged uploads do NOT persist across app
    restarts. Cleared on import. `bundle/` is left alone so a previous
    bundle's compose-run still resolves user_<id> references."""
    if USER_STAGED_DIR.exists():
        shutil.rmtree(USER_STAGED_DIR, ignore_errors=True)
    USER_STAGED_DIR.mkdir(parents=True, exist_ok=True)


def _user_meta_path(dir_: Path) -> Path:
    return dir_ / "meta.json"


def _read_user_meta(dir_: Path) -> dict:
    p = _user_meta_path(dir_)
    if not p.exists():
        return {}
    try:
        return json_mod.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _write_user_meta(dir_: Path, meta: dict) -> None:
    _user_meta_path(dir_).write_text(
        json_mod.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8",
    )


def _safe_user_filename(raw: str) -> tuple[str, str] | None:
    """Return (safe_basename, ext_lower) or None for rejected filenames.

    Strips path components and leading dots so a malicious upload can't
    escape the staging dir.
    """
    if not raw:
        return None
    base = Path(raw).name.lstrip(".").strip()
    if not base:
        return None
    ext = Path(base).suffix.lower()
    if ext not in _USER_EXT_KIND:
        return None
    return base, ext


def _user_asset_image_dims(path: Path) -> tuple[int | None, int | None]:
    try:
        from PIL import Image  # type: ignore
        with Image.open(path) as im:
            return int(im.width), int(im.height)
    except Exception:
        return None, None


def _user_asset_svg_dims(path: Path) -> tuple[int | None, int | None]:
    """Pull width/height from an SVG's root element if declared. Returns
    (None, None) on parse failure or missing attrs (acceptable — the
    agent gets to know it's a vector either way)."""
    try:
        import xml.etree.ElementTree as ET
        root = ET.parse(path).getroot()
    except Exception:
        return None, None

    def _strip_unit(v: str | None) -> int | None:
        if not v:
            return None
        v = v.strip()
        for unit in ("px", "pt", "mm", "cm", "in"):
            if v.endswith(unit):
                v = v[: -len(unit)]
                break
        try:
            return int(float(v))
        except ValueError:
            return None

    w = _strip_unit(root.get("width"))
    h = _strip_unit(root.get("height"))
    if (w is None or h is None) and root.get("viewBox"):
        parts = root.get("viewBox").replace(",", " ").split()
        if len(parts) == 4:
            w = w or _strip_unit(parts[2])
            h = h or _strip_unit(parts[3])
    return w, h


def _make_low_res(src: Path, dst: Path, kind: str) -> None:
    """For raster: downsize so long side <= _USER_LOW_RES_LONG_SIDE and
    re-encode. For everything else (svg / xml): copy verbatim."""
    if kind != "image":
        shutil.copyfile(src, dst)
        return
    try:
        from PIL import Image  # type: ignore
        with Image.open(src) as im:
            im.load()
            long_side = max(im.width, im.height)
            if long_side > _USER_LOW_RES_LONG_SIDE:
                scale = _USER_LOW_RES_LONG_SIDE / long_side
                new_size = (max(1, int(im.width * scale)),
                            max(1, int(im.height * scale)))
                im = im.resize(new_size, Image.LANCZOS)
            save_kwargs: dict = {}
            if dst.suffix.lower() in (".jpg", ".jpeg"):
                im = im.convert("RGB")
                save_kwargs["quality"] = 78
                save_kwargs["optimize"] = True
            elif dst.suffix.lower() == ".png":
                save_kwargs["optimize"] = True
            im.save(dst, **save_kwargs)
    except Exception:
        # On any failure: fall back to the original. Bundle size hit is
        # acceptable; correctness matters more.
        shutil.copyfile(src, dst)


def _hash_file(p: Path) -> str:
    """SHA1 of file content; first 8 chars used as the stable id."""
    import hashlib
    h = hashlib.sha1()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def _stage_user_asset(f, meta: dict) -> dict:
    """Validate + persist one uploaded file under USER_STAGED_DIR.

    Returns a metadata dict for the saved file, or raises ValueError.
    Mutates `meta` in place with the new entry.
    """
    parsed = _safe_user_filename(f.filename or "")
    if parsed is None:
        raise ValueError(f"filename rejected (need {sorted(_USER_EXT_KIND)})")
    orig_name, ext = parsed
    kind = _USER_EXT_KIND[ext]

    # Stream into a temp file inside the staging dir so we can hash + size-
    # check without buffering the whole upload in RAM.
    tmp = USER_STAGED_DIR / f".incoming_{datetime.now().strftime('%H%M%S%f')}{ext}"
    f.save(str(tmp))
    try:
        size_bytes = tmp.stat().st_size
        if size_bytes > _USER_MAX_FILE_BYTES:
            raise ValueError(
                f"file exceeds per-file cap of "
                f"{_USER_MAX_FILE_BYTES // (1024 * 1024)} MB"
            )
        total = sum(
            (e.get("size_bytes") or 0) for e in meta.values()
        ) + size_bytes
        if total > _USER_MAX_TOTAL_BYTES:
            raise ValueError(
                f"total staged size would exceed "
                f"{_USER_MAX_TOTAL_BYTES // (1024 * 1024)} MB"
            )
        if len(meta) >= _USER_MAX_FILES:
            raise ValueError(f"already at {_USER_MAX_FILES}-file cap")
        sha8 = _hash_file(tmp)[:8]
        aid = f"asset_{sha8}"
        # Dedupe: same content uploaded twice keeps the first entry.
        if aid in meta:
            tmp.unlink(missing_ok=True)
            return meta[aid]

        dims_w, dims_h = (None, None)
        if kind == "image":
            dims_w, dims_h = _user_asset_image_dims(tmp)
        elif kind == "vector":
            dims_w, dims_h = _user_asset_svg_dims(tmp)

        final = USER_STAGED_DIR / f"{aid}{ext}"
        tmp.rename(final)
        entry = {
            "id": aid,
            "filename": orig_name,
            "ext": ext.lstrip("."),
            "kind": kind,
            "size_bytes": size_bytes,
            "width": dims_w,
            "height": dims_h,
            "added_at": datetime.now().isoformat(timespec="seconds"),
        }
        meta[aid] = entry
        return entry
    except Exception:
        tmp.unlink(missing_ok=True)
        raise


def _move_staged_to_bundle() -> dict:
    """On bundle generation: replace USER_BUNDLE_DIR with current staged
    contents. Old bundle assets are removed (per user spec: 'base copies
    should be deleted after zip is generated so we don't keep bloat')."""
    if USER_BUNDLE_DIR.exists():
        shutil.rmtree(USER_BUNDLE_DIR, ignore_errors=True)
    USER_BUNDLE_DIR.mkdir(parents=True, exist_ok=True)
    staged_meta = _read_user_meta(USER_STAGED_DIR)
    moved: dict = {}
    for aid, entry in staged_meta.items():
        src = USER_STAGED_DIR / f"{aid}.{entry['ext']}"
        if not src.exists():
            continue
        dst = USER_BUNDLE_DIR / f"{aid}.{entry['ext']}"
        shutil.move(str(src), str(dst))
        moved[aid] = entry
    _write_user_meta(USER_BUNDLE_DIR, moved)
    # Clear staged for the next round.
    shutil.rmtree(USER_STAGED_DIR, ignore_errors=True)
    USER_STAGED_DIR.mkdir(parents=True, exist_ok=True)
    return moved


@app.post("/api/user_assets")
def api_user_assets_upload():
    """Upload one or more user-supplied assets (multipart 'files')."""
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "no files (multipart field 'files')"}), 400
    USER_STAGED_DIR.mkdir(parents=True, exist_ok=True)
    meta = _read_user_meta(USER_STAGED_DIR)
    added: list[dict] = []
    errors: list[dict] = []
    for f in files:
        try:
            entry = _stage_user_asset(f, meta)
            added.append(entry)
        except ValueError as e:
            errors.append({"filename": f.filename or "?", "reason": str(e)})
        except Exception as e:
            errors.append({"filename": f.filename or "?",
                           "reason": f"unexpected: {type(e).__name__}: {e}"})
    _write_user_meta(USER_STAGED_DIR, meta)
    return jsonify({"added": added, "errors": errors,
                    "staged": list(meta.values())})


@app.get("/api/user_assets")
def api_user_assets_list():
    meta = _read_user_meta(USER_STAGED_DIR)
    return jsonify({
        "staged": list(meta.values()),
        "total_bytes": sum((e.get("size_bytes") or 0) for e in meta.values()),
        "limits": {
            "max_files": _USER_MAX_FILES,
            "max_file_bytes": _USER_MAX_FILE_BYTES,
            "max_total_bytes": _USER_MAX_TOTAL_BYTES,
            "allowed_exts": sorted(_USER_EXT_KIND.keys()),
        },
    })


@app.delete("/api/user_assets/<asset_id>")
def api_user_assets_delete(asset_id):
    meta = _read_user_meta(USER_STAGED_DIR)
    entry = meta.pop(asset_id, None)
    if entry is None:
        return jsonify({"error": "asset id not staged"}), 404
    bin_path = USER_STAGED_DIR / f"{asset_id}.{entry.get('ext', '')}"
    bin_path.unlink(missing_ok=True)
    _write_user_meta(USER_STAGED_DIR, meta)
    return jsonify({"ok": True, "removed": asset_id})


@app.get("/api/user_assets/<asset_id>/preview")
def api_user_assets_preview(asset_id):
    """Stream the binary for thumbnail display in the UI. Looks at staged
    first, then bundle (so post-zip-gen UI still shows the same images)."""
    for d in (USER_STAGED_DIR, USER_BUNDLE_DIR):
        meta = _read_user_meta(d)
        entry = meta.get(asset_id)
        if entry is None:
            continue
        bin_path = d / f"{asset_id}.{entry.get('ext', '')}"
        if bin_path.exists():
            return send_file(bin_path)
    abort(404, "user asset not found")


# Clear staged dir on import so we don't surface zombie files from a
# previous app session (per user preference: tmp-style staging).
_clear_user_staged_on_startup()


@app.post("/api/compose/run")
def api_compose_run():
    body = request.get_json(force=True) or {}
    plan = body.get("plan")
    if not isinstance(plan, list) or not plan:
        return jsonify({"error": "plan must be a non-empty JSON array"}), 400

    import tempfile

    with tempfile.TemporaryDirectory(prefix="pptx-compose-") as tmpdir:
        staging = Path(tmpdir) / "bundle"
        staging.mkdir(parents=True, exist_ok=True)
        _stage_compose_bundle(staging)
        plan_path = staging / "plan.json"
        plan_path.write_text(json_mod.dumps(plan, ensure_ascii=False), encoding="utf-8")
        out_path = staging / "out.pptx"
        try:
            result = subprocess.run(
                [sys.executable, "reader.py", "compose", "plan.json", "out.pptx"],
                cwd=str(staging),
                capture_output=True,
                text=True,
                timeout=60,
            )
        except subprocess.TimeoutExpired:
            return jsonify({"error": "compose timed out"}), 500
        if result.returncode != 0 or not out_path.exists():
            return jsonify({
                "error": "compose failed",
                "stdout": result.stdout,
                "stderr": result.stderr,
            }), 500
        try:
            summary = json_mod.loads(result.stdout)
        except Exception:
            summary = {"stdout": result.stdout}
        # Persist out.pptx outside the temp dir so we can stream it after
        # cleanup of the staging dir.
        persisted = WORKSPACE / "_compose_out"
        persisted.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        out_persisted = persisted / f"deck-{ts}.pptx"
        shutil.copyfile(out_path, out_persisted)

    return send_file(
        out_persisted,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=out_persisted.name,
        max_age=0,
    )


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------


INDEX_HTML = r"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>pptx-skill describe</title>
  <style>
    * { box-sizing: border-box; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
           margin: 0; height: 100vh; display: flex; color: #222; }
    .sidebar { width: 280px; border-right: 1px solid #ddd; background: #fafafa;
               overflow-y: auto; display: flex; flex-direction: column; }
    .sidebar header { padding: 12px 14px; border-bottom: 1px solid #ddd; }
    .tabs { display: flex; border-bottom: 1px solid #ddd; }
    .tabs button { flex: 1; padding: 8px; border: none; background: none;
                   cursor: pointer; font-weight: 500; border-bottom: 2px solid transparent;
                   font-size: 13px; }
    .tabs button.active { border-bottom-color: #0066cc; color: #0066cc; }
    .filter-row { padding: 8px 14px; font-size: 12px; color: #555;
                  display: flex; gap: 6px; align-items: center; }
    .item-list { list-style: none; padding: 0; margin: 0; flex: 1; overflow-y: auto; }
    .item-list li { padding: 8px 14px; cursor: pointer; font-size: 12px;
                    display: flex; justify-content: space-between; align-items: center;
                    border-bottom: 1px solid #eee; }
    .item-list li:hover { background: #eef4ff; }
    .item-list li.active { background: #d8e8ff; }
    .pill { font-size: 9px; padding: 2px 6px; border-radius: 8px;
            text-transform: uppercase; letter-spacing: 0.4px; font-weight: 600; }
    .pill.pending { background: #ffeaa7; color: #8c6900; }
    .pill.done { background: #c8e6c9; color: #1b5e20; }
    .pill.locked { background: #d1c4e9; color: #311b92; }

    .preview { flex: 1; background: #1c1c1c; display: flex; align-items: center;
               justify-content: center; padding: 20px; min-width: 0; }
    .preview img { max-width: 100%; max-height: 100%;
                   box-shadow: 0 4px 30px rgba(0,0,0,0.4); background: white; }
    .preview .empty { color: #666; font-size: 13px; }

    .panel { width: 400px; border-left: 1px solid #ddd; padding: 16px;
             overflow-y: auto; background: white; }
    .panel h2 { margin: 0 0 4px; font-size: 15px; word-break: break-all; }
    .panel .sub { color: #888; font-size: 11px; margin-bottom: 12px;
                  word-break: break-all; }
    .panel label { display: block; font-size: 11px; font-weight: 700;
                   margin: 12px 0 4px; color: #555; text-transform: uppercase;
                   letter-spacing: 0.4px; }
    .panel label .hint { font-weight: 400; color: #999; text-transform: none;
                         letter-spacing: 0; margin-left: 6px; }
    .panel input, .panel select, .panel textarea {
      width: 100%; padding: 6px 8px; border: 1px solid #ccc; border-radius: 4px;
      font-size: 13px; font-family: inherit; background: white;
    }
    .panel textarea { resize: vertical; min-height: 50px; }
    .checks { display: flex; flex-wrap: wrap; gap: 4px; }
    .checks label { display: inline-flex; align-items: center; gap: 4px;
                    font-weight: normal; font-size: 11px; padding: 3px 8px;
                    border: 1px solid #ccc; border-radius: 12px; cursor: pointer;
                    background: #fff; margin: 0; text-transform: none;
                    letter-spacing: 0; }
    .checks input { width: auto; margin: 0; }
    .checks label:has(input:checked) { background: #d8e8ff; border-color: #0066cc;
                                        color: #003e7e; font-weight: 600; }

    .btnrow { display: flex; gap: 8px; margin-top: 16px; }
    button.primary { background: #0066cc; color: white; border: none;
                     padding: 8px 14px; border-radius: 4px; cursor: pointer;
                     font-weight: 500; font-size: 13px; }
    button.primary:hover { background: #0052a3; }
    button.ghost { background: white; color: #333; border: 1px solid #ccc;
                   padding: 8px 14px; border-radius: 4px; cursor: pointer;
                   font-size: 13px; }
    button.ghost:hover { background: #f0f0f0; }

    .mode-toggle { display: flex; gap: 0; border: 1px solid #ccc;
                   border-radius: 4px; overflow: hidden; margin: 12px 0; }
    .mode-toggle button { flex: 1; padding: 6px 10px; border: none;
                          background: white; cursor: pointer; font-size: 12px;
                          color: #555; font-weight: 500; }
    .mode-toggle button.active { background: #0066cc; color: white;
                                  font-weight: 600; }
    .mode-toggle button + button { border-left: 1px solid #ccc; }

    .paste-full textarea { width: 100%; min-height: 280px;
                            font-family: ui-monospace, Menlo, monospace;
                            font-size: 12px; padding: 10px; border: 1px solid #ccc;
                            border-radius: 4px; resize: vertical; }
    .paste-hint { font-size: 11px; color: #888; margin: 4px 0 8px; }
    .errors { background: #ffe9e9; color: #a00; padding: 8px;
              border-radius: 4px; margin-top: 10px; font-size: 12px;
              border: 1px solid #f5b8b8; }
    .errors ul { margin: 4px 0 0 16px; padding: 0; }
    .ok { color: #1b5e20; font-size: 12px; margin-top: 8px;
          background: #e8f5e9; padding: 8px; border-radius: 4px;
          border: 1px solid #c8e6c9; }
    .placeholder { color: #888; padding: 40px 16px; text-align: center;
                   font-size: 13px; }

    .batch-thumbs { padding: 20px; display: grid; align-content: start;
                     grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
                     gap: 12px; overflow-y: auto; max-height: 100%; width: 100%; }
    .batch-thumbs .bt { background: white; border-radius: 4px; padding: 6px;
                         box-shadow: 0 1px 4px rgba(0,0,0,0.25); text-align: center;
                         min-width: 0; }
    .batch-thumbs .bt img { width: 100%; height: 100px; object-fit: contain;
                             background: #f8f8f8; border-radius: 2px; display: block; }
    .batch-thumbs .bt .lbl { font-size: 11px; color: #555; margin-top: 4px;
                              font-family: ui-monospace, Menlo, monospace;
                              white-space: nowrap; overflow: hidden;
                              text-overflow: ellipsis; }
    .batch-thumbs-empty { color: #666; padding: 40px 20px; text-align: center;
                           font-size: 13px; }

    .view-toggle { padding: 8px 14px; border-bottom: 1px solid #ddd;
                   display: flex; gap: 6px; }
    .view-toggle button { flex: 1; padding: 6px 8px; border: 1px solid #ccc;
                          background: white; border-radius: 4px;
                          font-size: 12px; cursor: pointer; }
    .view-toggle button.active { background: #0066cc; color: white;
                                  border-color: #0066cc; font-weight: 600; }

    .batch-view { flex: 1; padding: 24px 32px; overflow-y: auto;
                  background: #f7f7f7; }
    .batch-view .card { background: white; border: 1px solid #e0e0e0;
                         border-radius: 6px; padding: 20px; margin-bottom: 16px; }
    .batch-view h2 { margin-top: 0; }
    .batch-view label.row { display: block; margin-bottom: 10px;
                             font-size: 12px; font-weight: 600; color: #555; }
    .batch-view .inline { display: flex; gap: 16px; align-items: center;
                           margin-bottom: 16px; }
    .batch-view .inline label { font-weight: 600; font-size: 12px; color: #555; }
    .batch-view input[type=number], .batch-view select {
      padding: 6px 8px; border: 1px solid #ccc; border-radius: 4px;
      font-size: 13px; }
    .batch-view textarea { width: 100%; min-height: 200px; padding: 10px;
                            font-family: ui-monospace, Menlo, monospace;
                            font-size: 12px; border: 1px solid #ccc;
                            border-radius: 4px; }
    .batch-view .results { font-size: 12px; margin-top: 12px; }
    .batch-view .results table { width: 100%; border-collapse: collapse; }
    .batch-view .results th, .batch-view .results td {
      text-align: left; padding: 6px 8px; border-bottom: 1px solid #eee; }
    .batch-view .results th { background: #fafafa; font-weight: 600;
                               font-size: 11px; text-transform: uppercase;
                               letter-spacing: 0.4px; color: #555; }
    .batch-view code { background: #eef; padding: 1px 4px; border-radius: 3px;
                        font-size: 11px; }
    .recent-batch { font-size: 12px; color: #555; margin: 8px 0; }

    .ingest-row { padding: 8px 14px; border-bottom: 1px solid #ddd;
                  background: #fafafa; display: flex; align-items: center;
                  gap: 8px; flex-wrap: wrap; }
    .ingest-row button { font-size: 12px; padding: 4px 10px;
                          border: 1px solid #0066cc; background: white;
                          color: #0066cc; border-radius: 4px; cursor: pointer; }
    .ingest-row button:hover { background: #eef4ff; }
    .ingest-row button:disabled { opacity: 0.5; cursor: not-allowed; }
    .ingest-row .ingest-msg { font-size: 11px; color: #555; flex: 1;
                               min-width: 0; overflow: hidden;
                               text-overflow: ellipsis; white-space: nowrap; }
    .ingest-row .ingest-msg.ok { color: #1b5e20; }
    .ingest-row .ingest-msg.err { color: #a00; white-space: normal; }
  </style>
</head>
<body>
  <aside class="sidebar">
    <header>
      <strong>pptx-skill</strong>
      <span id="counts" style="color:#888;font-size:11px;margin-left:6px;"></span>
      <a href="/compose" style="float:right;font-size:11px;color:#0066cc;
         text-decoration:none;">Compose →</a>
    </header>
    <div class="ingest-row">
      <input type="file" id="ingestFile" accept=".pptx"
             style="display:none;" />
      <button id="ingestBtn" type="button">+ Ingest .pptx</button>
      <span class="ingest-msg" id="ingestMsg"
            title="Upload a .pptx to ingest as a new deck. Re-uploads of an existing deck are rejected — delete its workspace/decks/&lt;name&gt;/ dir first.">
        Add a new deck to the workspace
      </span>
    </div>
    <div class="view-toggle">
      <button data-view="items" class="active">Single</button>
      <button data-view="batch">Bulk</button>
    </div>
    <div class="tabs" id="kindTabs">
      <button data-tab="slides" class="active">Slides</button>
      <button data-tab="assets">Assets</button>
    </div>
    <div class="filter-row" id="filterRow">
      <label><input type="checkbox" id="hideDone" checked> Hide done</label>
    </div>
    <ul class="item-list" id="list"></ul>
  </aside>

  <main class="preview" id="preview">
    <span class="empty">Pick an item from the sidebar</span>
  </main>

  <section class="batch-view" id="batchView" hidden>
    <div class="card">
      <h2>Bulk describe</h2>
      <p style="color:#666;font-size:13px;margin-top:0;">
        Package the next N pending items into a zip you can hand to a vision
        LLM. The model returns one JSON object; paste it back below to apply
        all descriptions at once.
      </p>
      <div class="inline">
        <label>Kind:
          <select id="batchKind">
            <option value="asset">Assets</option>
            <option value="slide">Slides</option>
            <option value="slide_with_assets">Slides + their assets (bundled)</option>
          </select>
        </label>
        <label>Count:
          <input type="number" id="batchCount" min="1" max="20" value="10">
        </label>
        <button class="primary" id="batchGenBtn">Generate batch</button>
        <span id="pendingHint" style="color:#888;font-size:12px;"></span>
      </div>
      <div class="recent-batch" id="recentBatch"></div>
    </div>

    <div class="card">
      <h2>Apply LLM response</h2>
      <p style="color:#666;font-size:12px;">
        Paste the JSON object the LLM returned. Items are matched by id
        (<code>"01"</code>, <code>"02"</code>, …) against the selected
        batch's saved manifest. Each entry is validated independently;
        entries that pass auto-promote to <code>done</code>.
      </p>
      <div class="inline" style="margin-bottom:12px;">
        <label>Target batch:
          <select id="batchSelect"></select>
        </label>
        <button class="ghost" id="batchRefreshBtn">↻</button>
        <span id="batchTargetLabel" style="color:#888;font-size:12px;"></span>
      </div>
      <textarea id="batchYaml" placeholder='{&#10;  "01": { "kind": "photo", "subject": "..." },&#10;  "02": { ... }&#10;}'></textarea>
      <div class="btnrow" style="margin-top:10px;">
        <button class="primary" id="batchApplyBtn">Apply batch</button>
      </div>
      <div class="results" id="batchResults"></div>
    </div>
  </section>

  <section class="panel" id="panel">
    <div class="placeholder" id="placeholder">Select an item to begin.</div>
    <div id="formWrap" hidden>
      <h2 id="itemTitle"></h2>
      <div class="sub" id="itemSub"></div>

      <div class="btnrow">
        <button class="ghost" id="copyPrompt">Copy describe prompt</button>
      </div>

      <div class="mode-toggle" role="tablist">
        <button data-mode="form" class="active">Form</button>
        <button data-mode="paste">Paste YAML</button>
      </div>

      <div id="formMode">
        <form id="form" onsubmit="return false;"></form>
      </div>

      <div id="pasteMode" class="paste-full" hidden>
        <div class="paste-hint">
          Paste the LLM's YAML response. Existing values shown so you can edit
          or replace entirely. Save promotes to <code>done</code> if valid.
        </div>
        <textarea id="pasteFull" spellcheck="false"></textarea>
      </div>

      <div id="msg"></div>
      <div class="btnrow">
        <button class="primary" id="saveBtn">Save → Next</button>
        <button class="ghost" id="saveOnly">Save</button>
      </div>
    </div>
  </section>

<script>
// Controlled vocab — fetched from /api/vocab on load (single source:
// authoring/schemas/vocab.yaml). Populated by `loadVocab()` before any
// form is built; do not edit inline.
let SLIDE_FEEL = [];
let SLIDE_TAGS = [];
let ASSET_KIND = [];
let ASSET_FEEL = [];
let ASSET_COMP = [];
let ASSET_TAGS = [];

async function loadVocab() {
  const r = await fetch("/api/vocab");
  if (!r.ok) throw new Error("vocab load failed");
  const v = await r.json();
  SLIDE_FEEL = v.slide.feel;
  SLIDE_TAGS = v.slide.suitable_for;
  ASSET_KIND = v.asset.kind;
  ASSET_FEEL = v.asset.feel;
  ASSET_COMP = v.asset.composition;
  ASSET_TAGS = v.asset.suitable_for;
}

let activeTab = "slides";
let items = {slides: [], assets: []};
let current = null;
let mode = localStorage.getItem("describe.mode") || "form";
let view = localStorage.getItem("describe.view") || "items";
let currentBatchId = localStorage.getItem("describe.batchId") || null;

function el(tag, attrs, ...kids) {
  const e = document.createElement(tag);
  attrs = attrs || {};
  for (const k in attrs) {
    if (k === "checked" || k === "selected") { if (attrs[k]) e[k] = true; }
    else if (k === "html") e.innerHTML = attrs[k];
    else e.setAttribute(k, attrs[k]);
  }
  kids.forEach(k => e.append(k));
  return e;
}

async function loadItems() {
  const r = await fetch("/api/items");
  items = await r.json();
  renderCounts();
  renderList();
}

function renderCounts() {
  const sd = items.slides.filter(i => i.status === "done").length;
  const ad = items.assets.filter(i => i.status === "done").length;
  document.getElementById("counts").textContent =
    `slides ${sd}/${items.slides.length} · assets ${ad}/${items.assets.length}`;
  refreshPendingHint();
}

function renderList() {
  const list = document.getElementById("list");
  list.innerHTML = "";
  const hideDone = document.getElementById("hideDone").checked;
  const pool = items[activeTab];
  pool.forEach(it => {
    if (hideDone && it.status === "done") return;
    const li = el("li", {});
    if (current && current.yaml === it.yaml) li.classList.add("active");
    li.addEventListener("click", () => loadItem(it.yaml));
    li.append(el("span", {}, it.id));
    li.append(el("span", {class: "pill " + it.status}, it.status));
    list.append(li);
  });
}

async function loadItem(yamlRel) {
  const r = await fetch("/api/item?yaml=" + encodeURIComponent(yamlRel));
  const item = await r.json();
  current = item;
  document.getElementById("preview").innerHTML =
    `<img src="/preview?yaml=${encodeURIComponent(yamlRel)}&t=${Date.now()}" alt="preview">`;
  document.getElementById("placeholder").hidden = true;
  document.getElementById("formWrap").hidden = false;
  document.getElementById("itemTitle").textContent = item.data.id || yamlRel;
  document.getElementById("itemSub").textContent = yamlRel;
  buildForm(item);
  document.getElementById("pasteFull").value = item.yaml_text || "";
  renderList();
}

function applyMode() {
  document.querySelectorAll(".mode-toggle button").forEach(b => {
    b.classList.toggle("active", b.dataset.mode === mode);
  });
  document.getElementById("formMode").hidden = mode !== "form";
  document.getElementById("pasteMode").hidden = mode !== "paste";
}

function setMode(next) {
  mode = next;
  localStorage.setItem("describe.mode", mode);
  applyMode();
}

function buildForm(item) {
  const f = document.getElementById("form");
  f.innerHTML = "";
  const d = item.data || {};
  if (item.kind === "slide") {
    addText(f, "intent", d.intent || "", "one sentence, <20 words");
    addSelect(f, "feel", d.feel || "", SLIDE_FEEL);
    addChips(f, "suitable_for", d.suitable_for || [], SLIDE_TAGS);
    addTextarea(f, "notes", d.notes || "", "human reviewer note");
    addTextarea(f, "interpretation", d.interpretation || "",
      "model's speculative observations — info only, not filterable");
  } else {
    addSelect(f, "kind", d.kind || "", ASSET_KIND);
    addText(f, "subject", d.subject || "", "neutral, <25 words");
    addText(f, "depicts", d.depicts || "", "the concept — 1-5 words; empty for decorative");
    addSelect(f, "feel", d.feel || "", ASSET_FEEL);
    addSelect(f, "composition", d.composition || "", ASSET_COMP);
    addText(f, "colors", (d.colors || []).join(", "), "1-3 words, comma-separated");
    addText(f, "scope", (d.scope || []).join(", "),
      "comma-separated; e.g. 'client:acme-bank, industry:finance' or 'generic'");
    addChips(f, "suitable_for", d.suitable_for || [], ASSET_TAGS);
    addTextarea(f, "notes", d.notes || "", "human reviewer note");
    addTextarea(f, "interpretation", d.interpretation || "",
      "model's speculative observations — info only, not filterable");
  }
  document.getElementById("msg").innerHTML = "";
}

function addText(parent, name, value, hint) {
  const lbl = el("label", {for: name}, name);
  if (hint) lbl.append(el("span", {class: "hint"}, hint));
  parent.append(lbl);
  parent.append(el("input", {type: "text", name, value, id: name}));
}
function addTextarea(parent, name, value, hint) {
  const lbl = el("label", {for: name}, name);
  if (hint) lbl.append(el("span", {class: "hint"}, hint));
  parent.append(lbl);
  parent.append(el("textarea", {name, id: name, rows: "3"}, value));
}
function addSelect(parent, name, value, options) {
  parent.append(el("label", {for: name}, name));
  const sel = el("select", {name, id: name});
  sel.append(el("option", {value: ""}, "—"));
  options.forEach(o => sel.append(el("option", {value: o, selected: o === value}, o)));
  parent.append(sel);
}
function addChips(parent, name, values, options) {
  parent.append(el("label", {}, name));
  const wrap = el("div", {class: "checks", id: name});
  options.forEach(o => {
    const lbl = el("label", {},
      el("input", {type: "checkbox", value: o, checked: values.includes(o)}),
      o
    );
    wrap.append(lbl);
  });
  parent.append(wrap);
}

function gatherForm() {
  if (!current) return {};
  const out = {};
  if (current.kind === "slide") {
    out.intent = document.getElementById("intent").value.trim();
    out.feel = document.getElementById("feel").value;
    out.suitable_for = chipValues("suitable_for");
    out.notes = document.getElementById("notes").value.trim();
    out.interpretation = document.getElementById("interpretation").value.trim();
  } else {
    out.kind = document.getElementById("kind").value;
    out.subject = document.getElementById("subject").value.trim();
    out.depicts = document.getElementById("depicts").value.trim();
    out.feel = document.getElementById("feel").value;
    out.composition = document.getElementById("composition").value;
    out.colors = document.getElementById("colors").value
      .split(",").map(s => s.trim()).filter(Boolean);
    out.scope = document.getElementById("scope").value
      .split(",").map(s => s.trim()).filter(Boolean);
    out.suitable_for = chipValues("suitable_for");
    out.notes = document.getElementById("notes").value.trim();
    out.interpretation = document.getElementById("interpretation").value.trim();
  }
  return out;
}

function chipValues(name) {
  return [...document.querySelectorAll("#" + name + " input:checked")].map(c => c.value);
}

async function save(advance) {
  if (!current) return;
  let r;
  if (mode === "paste") {
    const text = document.getElementById("pasteFull").value;
    r = await fetch("/api/save-raw", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({yaml: current.yaml, text}),
    });
  } else {
    const fields = gatherForm();
    r = await fetch("/api/save", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({yaml: current.yaml, fields}),
    });
  }
  const result = await r.json();
  const msg = document.getElementById("msg");
  if (result.errors && result.errors.length) {
    msg.innerHTML = '<div class="errors"><strong>Validation errors</strong><ul>'
      + result.errors.map(e => "<li>" + escapeHtml(e) + "</li>").join("")
      + '</ul></div>';
    return;
  }
  msg.innerHTML = '<div class="ok">Saved — status: ' + escapeHtml(result.status || "") + '</div>';
  const poolKey = current.kind === "slide" ? "slides" : "assets";
  const it = items[poolKey].find(i => i.yaml === current.yaml);
  if (it) it.status = result.status;
  renderCounts();
  renderList();
  if (advance) {
    const next = items[poolKey].find(i => i.status === "pending");
    if (next) loadItem(next.yaml);
    else msg.innerHTML += '<div class="ok" style="margin-top:6px;">No more pending in this tab 🎉</div>';
  }
}

async function copyPrompt() {
  if (!current) return;
  const r = await fetch("/api/prompt?kind=" + current.kind);
  const j = await r.json();
  try {
    await navigator.clipboard.writeText(j.text);
    const btn = document.getElementById("copyPrompt");
    const orig = btn.textContent;
    btn.textContent = "Copied ✓";
    setTimeout(() => (btn.textContent = orig), 1500);
  } catch (e) {
    alert("Could not copy to clipboard:\n" + e.message + "\n\nPrompt:\n\n" + j.text);
  }
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])
  );
}

document.querySelectorAll(".tabs button").forEach(b => {
  b.addEventListener("click", () => {
    activeTab = b.dataset.tab;
    document.querySelectorAll(".tabs button").forEach(x => x.classList.toggle("active", x === b));
    renderList();
  });
});
document.querySelectorAll(".view-toggle button").forEach(b => {
  b.addEventListener("click", () => setView(b.dataset.view));
});

function applyView() {
  document.querySelectorAll(".view-toggle button").forEach(b => {
    b.classList.toggle("active", b.dataset.view === view);
  });
  const inItems = view === "items";
  document.getElementById("kindTabs").style.display = inItems ? "" : "none";
  document.getElementById("filterRow").style.display = inItems ? "" : "none";
  document.getElementById("list").style.display = inItems ? "" : "none";
  document.getElementById("panel").hidden = !inItems;
  document.getElementById("batchView").hidden = inItems;
  if (inItems) {
    const pane = document.getElementById("preview");
    if (current) {
      pane.innerHTML =
        `<img src="/preview?yaml=${encodeURIComponent(current.yaml)}&t=${Date.now()}" alt="preview">`;
    } else {
      pane.innerHTML = '<span class="empty">Pick an item from the sidebar</span>';
    }
  } else {
    loadBatches().then(() => renderBatchThumbnails(currentBatchId));
  }
}

async function renderBatchThumbnails(batchId) {
  const pane = document.getElementById("preview");
  if (!batchId) {
    pane.innerHTML =
      '<div class="batch-thumbs-empty">Generate or select a batch to see its items.</div>';
    return;
  }
  pane.innerHTML = '<div class="batch-thumbs-empty">loading…</div>';
  try {
    const r = await fetch("/api/batch/" + encodeURIComponent(batchId));
    if (!r.ok) throw new Error("batch fetch failed");
    const m = await r.json();
    const items = m.items || {};
    if (!Object.keys(items).length) {
      pane.innerHTML = '<div class="batch-thumbs-empty">Batch is empty.</div>';
      return;
    }
    const cards = Object.entries(items).map(([k, v]) => {
      // slide_with_assets entries are {slide, assets}; render the slide
      // preview and a small count of bundled assets.
      if (v && typeof v === "object" && v.slide) {
        const nAssets = Object.keys(v.assets || {}).length;
        const slideRel = v.slide;
        return `<div class="bt" title="${escapeHtml(slideRel)}">
          <img src="/preview?yaml=${encodeURIComponent(slideRel)}" alt="${escapeHtml(k)}" loading="lazy">
          <div class="lbl">${escapeHtml(k)} · ${escapeHtml(slideRel.split("/").pop().replace(/\.yaml$/, ""))} <em>(+${nAssets})</em></div>
         </div>`;
      }
      return `<div class="bt" title="${escapeHtml(v)}">
        <img src="/preview?yaml=${encodeURIComponent(v)}" alt="${escapeHtml(k)}" loading="lazy">
        <div class="lbl">${escapeHtml(k)} · ${escapeHtml(v.split("/").pop().replace(/\.yaml$/, ""))}</div>
       </div>`;
    }).join("");
    pane.innerHTML = `<div class="batch-thumbs">${cards}</div>`;
  } catch (e) {
    pane.innerHTML =
      '<div class="batch-thumbs-empty">Could not load batch (' + escapeHtml(e.message) + ').</div>';
  }
}
function setView(next) {
  view = next;
  localStorage.setItem("describe.view", view);
  applyView();
}

function refreshBatchLabel() {
  const lbl = document.getElementById("batchTargetLabel");
  if (currentBatchId) {
    lbl.textContent = "";
  } else {
    lbl.textContent = "(generate one above)";
  }
}

async function loadBatches(selectId) {
  const sel = document.getElementById("batchSelect");
  const r = await fetch("/api/batches");
  const j = await r.json();
  sel.innerHTML = "";
  if (!j.batches.length) {
    const opt = document.createElement("option");
    opt.value = ""; opt.textContent = "— no batches yet —";
    sel.append(opt);
    currentBatchId = null;
    localStorage.removeItem("describe.batchId");
    refreshBatchLabel();
    return;
  }
  j.batches.forEach(b => {
    const opt = document.createElement("option");
    opt.value = b.batch_id;
    opt.textContent = `${b.batch_id} — ${b.kind} × ${b.count}`;
    sel.append(opt);
  });
  const desired = selectId || currentBatchId || j.batches[0].batch_id;
  const exists = j.batches.some(b => b.batch_id === desired);
  sel.value = exists ? desired : j.batches[0].batch_id;
  currentBatchId = sel.value;
  localStorage.setItem("describe.batchId", currentBatchId);
  refreshBatchLabel();
}

function _shortPath(p) { return p.split("/").slice(-2).join("/"); }

function _skippedHtml(skipped) {
  if (!skipped || !skipped.length) return "";
  const rows = skipped.map(s =>
    `<li><code>${escapeHtml(_shortPath(s.yaml))}</code>: ${escapeHtml(s.reason)}</li>`
  ).join("");
  return `<div style="background:#fff8e1;border:1px solid #ffd980;padding:8px;
    border-radius:4px;margin-top:8px;font-size:12px;">
    <strong>${skipped.length} skipped</strong> (not in zip)
    <ul style="margin:4px 0 0 16px;">${rows}</ul></div>`;
}

function refreshPendingHint() {
  const kind = document.getElementById("batchKind").value;
  const pool = kind === "asset" ? items.assets : items.slides;
  const n = pool.filter(i => i.status === "pending").length;
  const hint = document.getElementById("pendingHint");
  const label = (
    kind === "asset" ? "asset" :
    kind === "slide" ? "slide" :
    "slide bundle"
  );
  hint.textContent = `(${n} pending ${label}${n === 1 ? "" : "s"})`;
  hint.style.color = n === 0 ? "#a00" : "#888";
  document.getElementById("batchGenBtn").disabled = n === 0;
}

async function batchGenerate() {
  const kind = document.getElementById("batchKind").value;
  const count = parseInt(document.getElementById("batchCount").value, 10) || 10;
  const recent = document.getElementById("recentBatch");
  recent.innerHTML = '<em style="color:#888;">generating…</em>';

  const r = await fetch("/api/batch/create", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({kind, count}),
  });

  if (!r.ok) {
    const e = await r.json().catch(() => ({error: "unknown error"}));
    recent.innerHTML =
      `<div class="errors"><strong>${escapeHtml(e.error || r.statusText)}</strong></div>`
      + _skippedHtml(e.skipped);
    return;
  }

  const j = await r.json();
  currentBatchId = j.batch_id;
  localStorage.setItem("describe.batchId", currentBatchId);

  const itemsList = Object.entries(j.items).map(([k, v]) => {
    // slide_with_assets entries are nested {slide, assets}; others are strings.
    if (v && typeof v === "object" && v.slide) {
      const nAssets = Object.keys(v.assets || {}).length;
      return `<code>${k}</code> → ${escapeHtml(_shortPath(v.slide))} <em>(+${nAssets} asset${nAssets === 1 ? "" : "s"})</em>`;
    }
    return `<code>${k}</code> → ${escapeHtml(_shortPath(v))}`;
  }).join("<br>");

  const summary = (j.count < (j.requested || j.count))
    ? `${j.count} of ${j.requested} requested`
    : `${j.count} ${j.kind}(s)`;

  recent.innerHTML =
    `<strong>Batch ${j.batch_id}</strong> — ${summary}. ` +
    `<a href="${j.download_url}" download>Download zip</a><br>` +
    `<details style="margin-top:6px;"><summary>Items</summary>${itemsList}</details>`
    + _skippedHtml(j.skipped);

  await loadBatches(currentBatchId);
  refreshPendingHint();
  renderBatchThumbnails(currentBatchId);
  window.location.href = j.download_url;
}

async function batchApply() {
  if (!currentBatchId) {
    alert("Generate a batch first.");
    return;
  }
  const text = document.getElementById("batchYaml").value;
  if (!text.trim()) { alert("Paste the LLM's JSON first."); return; }
  const r = await fetch("/api/batch/" + currentBatchId + "/apply", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({text}),
  });
  const j = await r.json();
  if (j.error) {
    document.getElementById("batchResults").innerHTML =
      `<div class="errors">${escapeHtml(j.error)}</div>`;
    return;
  }
  const rows = j.results.map(rr => {
    const pillCls = rr.status === "done" ? "done" :
                    rr.status === "pending" ? "pending" : "locked";
    const errs = (rr.errors && rr.errors.length)
      ? "<br><small style='color:#a00;'>" + rr.errors.map(escapeHtml).join("; ") + "</small>"
      : "";
    return `<tr>
      <td><code>${escapeHtml(rr.id)}</code></td>
      <td>${escapeHtml(rr.yaml.split("/").slice(-2).join("/"))}</td>
      <td><span class="pill ${pillCls}">${escapeHtml(rr.status)}</span>${errs}</td>
    </tr>`;
  }).join("");
  let diag = "";
  if (j.found_keys) {
    const matched = j.matched || 0;
    diag = `<div style="margin-bottom:8px;font-size:12px;color:#555;">
      <strong>${matched}/${j.results.length}</strong> matched.
      Top-level keys found in your YAML:
      ${j.found_keys.length
        ? j.found_keys.map(k => "<code>" + escapeHtml(k) + "</code>").join(" ")
        : "<em>none</em>"}
    </div>`;
  }
  document.getElementById("batchResults").innerHTML = diag +
    `<table><thead><tr><th>Id</th><th>Target</th><th>Result</th></tr></thead>
     <tbody>${rows}</tbody></table>`;
  // Refresh sidebar counts so user sees the progress reflected
  loadItems();
}

document.getElementById("batchGenBtn").addEventListener("click", batchGenerate);
document.getElementById("batchApplyBtn").addEventListener("click", batchApply);
document.getElementById("batchKind").addEventListener("change", refreshPendingHint);
document.getElementById("batchSelect").addEventListener("change", e => {
  currentBatchId = e.target.value || null;
  if (currentBatchId) localStorage.setItem("describe.batchId", currentBatchId);
  else localStorage.removeItem("describe.batchId");
  refreshBatchLabel();
  if (view === "batch") renderBatchThumbnails(currentBatchId);
});
document.getElementById("batchRefreshBtn").addEventListener("click", async () => {
  await loadBatches();
  if (view === "batch") renderBatchThumbnails(currentBatchId);
});
document.getElementById("hideDone").addEventListener("change", renderList);
document.getElementById("saveBtn").addEventListener("click", () => save(true));
document.getElementById("saveOnly").addEventListener("click", () => save(false));
document.getElementById("copyPrompt").addEventListener("click", copyPrompt);
document.querySelectorAll(".mode-toggle button").forEach(b => {
  b.addEventListener("click", () => setMode(b.dataset.mode));
});

// --- Ingest .pptx upload ---------------------------------------------------
const ingestBtn = document.getElementById("ingestBtn");
const ingestFile = document.getElementById("ingestFile");
const ingestMsg = document.getElementById("ingestMsg");

function setIngestMsg(text, tone) {
  ingestMsg.textContent = text;
  ingestMsg.classList.remove("ok", "err");
  if (tone) ingestMsg.classList.add(tone);
}

ingestBtn.addEventListener("click", () => ingestFile.click());
ingestFile.addEventListener("change", async () => {
  const f = ingestFile.files && ingestFile.files[0];
  if (!f) return;
  setIngestMsg("Uploading " + f.name + " …", null);
  ingestBtn.disabled = true;
  const fd = new FormData();
  fd.append("pptx", f);
  try {
    const r = await fetch("/api/ingest", { method: "POST", body: fd });
    const data = await r.json().catch(() => ({error: "non-JSON response"}));
    if (!r.ok) {
      setIngestMsg(data.error || ("HTTP " + r.status), "err");
    } else {
      setIngestMsg(
        "Ingested " + data.deck_stem + ": "
        + data.slides + " slides, "
        + data.pictures + " pictures, "
        + data.atoms + " atoms",
        "ok"
      );
      loadItems();  // refresh sidebar so new pending items appear
    }
  } catch (e) {
    setIngestMsg("Upload failed: " + e.message, "err");
  } finally {
    ingestBtn.disabled = false;
    ingestFile.value = "";  // allow re-uploading same filename
  }
});

applyMode();
applyView();
loadVocab().then(loadItems).catch(err => {
  document.getElementById("msg").innerHTML = "vocab load failed: " + err.message;
  loadItems();
});
</script>
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(INDEX_HTML)


COMPOSE_HTML = r"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>pptx-skill — compose</title>
  <style>
    * { box-sizing: border-box; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
           margin: 0; background: #f7f7f7; color: #222; }
    .top { background: white; border-bottom: 1px solid #ddd; padding: 12px 24px;
           display: flex; align-items: center; gap: 16px; }
    .top strong { font-size: 16px; }
    .top a { color: #0066cc; text-decoration: none; font-size: 13px; }
    .wrap { max-width: 880px; margin: 24px auto; padding: 0 24px; }
    .step { background: white; border: 1px solid #e0e0e0; border-radius: 6px;
            padding: 20px 24px; margin-bottom: 16px; }
    .step h2 { margin: 0 0 4px; font-size: 16px; }
    .step .desc { color: #666; font-size: 13px; margin-bottom: 14px; }
    .filter-block { margin-bottom: 14px; }
    .filter-block h3 { font-size: 12px; text-transform: uppercase;
                       letter-spacing: 0.5px; color: #555; margin: 0 0 6px;
                       font-weight: 600; }
    .chips { display: flex; flex-wrap: wrap; gap: 6px; }
    .chip { background: #f0f0f0; border: 1px solid #ddd; border-radius: 999px;
            padding: 4px 10px; font-size: 12px; cursor: pointer;
            user-select: none; }
    .chip.on { background: #0066cc; color: white; border-color: #0066cc; }
    .chip:hover { border-color: #0066cc; }
    .count-row { background: #eef4ff; border: 1px solid #c8dcf5; border-radius: 4px;
                  padding: 8px 12px; font-size: 13px; color: #234;
                  margin: 10px 0 0; }
    textarea { width: 100%; padding: 10px; border: 1px solid #ccc;
               border-radius: 4px; font-size: 13px; resize: vertical;
               font-family: ui-monospace, Menlo, monospace; }
    .brief-area { min-height: 90px; font-family: inherit; }
    .plan-area { min-height: 200px; }
    .actions { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 12px; }
    button { padding: 8px 14px; border: 1px solid #0066cc; background: #0066cc;
             color: white; border-radius: 4px; cursor: pointer; font-size: 13px;
             font-weight: 500; }
    button.ghost { background: white; color: #0066cc; }
    button:disabled { opacity: 0.5; cursor: not-allowed; }
    .msg { font-size: 12px; padding: 8px 10px; border-radius: 4px;
           margin-top: 10px; display: none; }
    .msg.ok { background: #e8f5e9; color: #1b5e20; border: 1px solid #c8e6c9;
              display: block; }
    .msg.err { background: #ffe9e9; color: #a00; border: 1px solid #f5b8b8;
               display: block; }
    .group-label { font-size: 13px; font-weight: 600; color: #333;
                   margin: 12px 0 6px; }
    pre.preview-text { background: #f8f8f8; border: 1px solid #ddd;
                       border-radius: 4px; padding: 10px; max-height: 220px;
                       overflow: auto; font-size: 11px; white-space: pre-wrap;
                       margin: 8px 0 0; }
    details.filter-details { border: 1px solid #e0e0e0; border-radius: 4px;
                             margin: 0 0 8px; background: #fcfcfc; }
    details.filter-details > summary { padding: 8px 12px; cursor: pointer;
                                        font-size: 13px; display: flex;
                                        align-items: center; gap: 10px;
                                        user-select: none; list-style: none; }
    details.filter-details > summary::-webkit-details-marker { display: none; }
    details.filter-details > summary::before { content: "▸"; font-size: 11px;
                                                color: #888; transition: transform 0.15s; }
    details.filter-details[open] > summary::before { transform: rotate(90deg);
                                                      display: inline-block; }
    details.filter-details > summary .fname { font-weight: 600; color: #333; }
    details.filter-details > summary .fcount { color: #888; font-size: 12px;
                                                margin-left: auto; }
    details.filter-details > summary .fbtns { display: flex; gap: 4px; }
    details.filter-details > summary .fbtn { font-size: 11px; color: #0066cc;
                                              cursor: pointer; padding: 2px 6px;
                                              border-radius: 3px; }
    details.filter-details > summary .fbtn:hover { background: #eef4ff; }
    details.filter-details .body { padding: 4px 12px 12px; }
    .preset-row { display: flex; gap: 8px; align-items: center;
                   margin-bottom: 10px; font-size: 13px; }
    .preset-row select { padding: 6px 8px; border: 1px solid #ccc;
                          border-radius: 4px; font-size: 13px; flex: 1; }
    .preset-row .small { background: white; color: #0066cc;
                          border: 1px solid #0066cc; padding: 6px 10px; }
    .brand-editor textarea { min-height: 140px; font-family: ui-monospace,
                              Menlo, monospace; font-size: 12px; }
    .brand-summary { font-size: 12px; color: #666; }
    .brand-summary code { background: #f0f0f0; padding: 1px 4px;
                           border-radius: 3px; font-size: 11px; }
  </style>
</head>
<body>
  <div class="top">
    <strong>pptx-skill</strong>
    <a href="/">← describe</a>
    <span style="color:#999;font-size:12px;margin-left:auto;" id="kbSummary"></span>
  </div>

  <div class="wrap">

    <details class="step brand-editor" id="brandStep">
      <summary style="cursor:pointer;list-style:none;">
        <h2 style="display:inline;">Brand rules <span class="brand-summary" id="brandSummary"></span></h2>
      </summary>
      <div class="desc" style="margin-top:8px;">
        Org-wide constraints (palette, voice, taboos). Auto-included in every
        prompt bundle. Edits save to <code class="brand-summary">authoring/brand.md</code>.
      </div>
      <textarea id="brand" placeholder="# Palette&#10;- Primary: ...&#10;# Voice&#10;- ..."></textarea>
      <div class="actions">
        <button id="saveBrand">Save brand.md</button>
      </div>
      <div class="msg" id="brandMsg"></div>
    </details>

    <div class="step">
      <h2>1. Filter the KB for this deck</h2>
      <div class="desc">
        Pick tags to narrow what the agent sees. Empty = no filter on that field.
        Within a field, multiple selections mean OR.
      </div>
      <div class="group-label">Templates</div>
      <div id="tplFilters"></div>
      <div class="group-label">Assets</div>
      <div id="astFilters"></div>
      <div class="count-row" id="countRow">matching: …</div>
    </div>

    <div class="step">
      <h2>2. Describe the deck you want</h2>
      <div class="desc">
        Topic, audience, length, tone. Goes into the bundle as <code>brief.md</code>.
      </div>
      <div class="preset-row">
        <span>Preset:</span>
        <select id="presetSelect"><option value="">— none —</option></select>
        <button class="small" id="loadPreset">Load</button>
        <button class="small" id="savePreset">Save as…</button>
        <button class="small" id="deletePreset">Delete</button>
      </div>
      <textarea id="brief" class="brief-area" placeholder="e.g. 4-slide thesis-defense summary for an academic committee. Formal feel. Include the swimlane diagram on the methodology slide."></textarea>
      <div class="actions">
        <button id="dlBundle">Download bundle (.zip)</button>
        <button id="copyText" class="ghost">Copy as text</button>
        <button id="showText" class="ghost">Preview text</button>
      </div>
      <div class="msg" id="bundleMsg"></div>
      <pre class="preview-text" id="previewText" hidden></pre>
    </div>

    <div class="step">
      <h2>3. Paste the agent's plan and compose</h2>
      <div class="desc">
        Paste the JSON array the LLM returns. We run <code>reader.py compose</code>
        and hand you the .pptx.
      </div>
      <textarea id="plan" class="plan-area" placeholder='[{"template":"…","slots":{"title":"…"}}]'></textarea>
      <div class="actions">
        <button id="runCompose">Compose deck (.pptx)</button>
      </div>
      <div class="msg" id="composeMsg"></div>
    </div>

  </div>

<script>
const tplFields = ["feel", "suitable_for"];
const astFields = ["kind", "feel", "composition", "suitable_for", "scope", "colors"];
const state = { templates: {}, assets: {} };

function fieldLabel(f) {
  return ({
    feel: "feel",
    suitable_for: "suitable for",
    kind: "kind",
    composition: "composition",
    scope: "scope",
    colors: "colors",
  })[f] || f;
}

function buildDimension(parent, scope, field, vals) {
  const det = document.createElement("details");
  det.className = "filter-details";
  det.dataset.scope = scope;
  det.dataset.field = field;
  const summary = document.createElement("summary");
  summary.innerHTML =
    `<span class="fname">${fieldLabel(field)}</span>` +
    `<span class="fcount" data-role="count">0 / ${vals.length}</span>` +
    `<span class="fbtns">` +
    `  <span class="fbtn" data-act="all">all</span>` +
    `  <span class="fbtn" data-act="clear">clear</span>` +
    `</span>`;
  det.appendChild(summary);
  const body = document.createElement("div");
  body.className = "body";
  const chips = document.createElement("div");
  chips.className = "chips";
  vals.forEach((v) => {
    const c = document.createElement("span");
    c.className = "chip";
    c.dataset.value = v;
    c.textContent = v;
    c.onclick = (e) => { e.stopPropagation(); toggle(scope, field, v, c, det); };
    chips.appendChild(c);
  });
  body.appendChild(chips);
  det.appendChild(body);
  summary.querySelectorAll(".fbtn").forEach((btn) => {
    btn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      const act = btn.dataset.act;
      state[scope][field] = [];
      if (act === "all") {
        vals.forEach((v) => state[scope][field].push(v));
      }
      chips.querySelectorAll(".chip").forEach((c) => {
        c.classList.toggle("on", state[scope][field].includes(c.dataset.value));
      });
      updateDimensionCount(det, scope, field, vals.length);
      refreshCount();
    };
  });
  updateDimensionCount(det, scope, field, vals.length);
  parent.appendChild(det);
}

function updateDimensionCount(det, scope, field, total) {
  const sel = (state[scope][field] || []).length;
  det.querySelector('[data-role="count"]').textContent = `${sel} / ${total}`;
}

function renderFilters() {
  const tplEl = document.getElementById("tplFilters");
  const astEl = document.getElementById("astFilters");
  tplEl.innerHTML = "";
  astEl.innerHTML = "";
  if (state.options.templates.options) {
    tplFields.forEach((f) => {
      const vals = state.options.templates.options[f] || [];
      if (vals.length) buildDimension(tplEl, "templates", f, vals);
    });
  }
  if (state.options.assets.options) {
    astFields.forEach((f) => {
      const vals = state.options.assets.options[f] || [];
      if (vals.length) buildDimension(astEl, "assets", f, vals);
    });
  }
  document.getElementById("kbSummary").textContent =
    `KB: ${state.options.templates.total} templates · ${state.options.assets.total} assets`;
}

function toggle(scope, field, value, el, det) {
  state[scope][field] = state[scope][field] || [];
  const i = state[scope][field].indexOf(value);
  if (i >= 0) {
    state[scope][field].splice(i, 1);
    el.classList.remove("on");
  } else {
    state[scope][field].push(value);
    el.classList.add("on");
  }
  if (det) {
    const total = det.querySelectorAll(".chip").length;
    updateDimensionCount(det, scope, field, total);
  }
  refreshCount();
}

async function refreshCount() {
  const r = await fetch("/api/compose/preview", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ filters: { templates: state.templates, assets: state.assets } }),
  });
  const j = await r.json();
  document.getElementById("countRow").textContent =
    `matching: ${j.templates} template(s), ${j.assets} asset(s)`;
}

function currentFilters() {
  return { templates: state.templates, assets: state.assets };
}

function showMsg(id, text, ok) {
  const el = document.getElementById(id);
  el.className = "msg " + (ok ? "ok" : "err");
  el.textContent = text;
}

document.getElementById("dlBundle").onclick = async () => {
  showMsg("bundleMsg", "building…", true);
  const r = await fetch("/api/compose/bundle", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ filters: currentFilters(), brief: document.getElementById("brief").value }),
  });
  if (!r.ok) { showMsg("bundleMsg", (await r.json()).error || "bundle failed", false); return; }
  const blob = await r.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const cd = r.headers.get("Content-Disposition") || "";
  const m = cd.match(/filename="?([^";]+)"?/);
  a.download = m ? m[1] : "prompt-bundle.zip";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  showMsg("bundleMsg", "downloaded " + a.download, true);
};

async function fetchFlat() {
  const r = await fetch("/api/compose/text", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ filters: currentFilters(), brief: document.getElementById("brief").value }),
  });
  return (await r.json()).text || "";
}

document.getElementById("copyText").onclick = async () => {
  showMsg("bundleMsg", "preparing…", true);
  const text = await fetchFlat();
  try {
    await navigator.clipboard.writeText(text);
    showMsg("bundleMsg", `copied ${text.length.toLocaleString()} chars to clipboard`, true);
  } catch (e) {
    showMsg("bundleMsg", "clipboard blocked — use Preview text and copy manually", false);
  }
};

document.getElementById("showText").onclick = async () => {
  const el = document.getElementById("previewText");
  if (!el.hidden) { el.hidden = true; return; }
  el.textContent = await fetchFlat();
  el.hidden = false;
};

document.getElementById("runCompose").onclick = async () => {
  const raw = document.getElementById("plan").value.trim();
  if (!raw) { showMsg("composeMsg", "paste a plan first", false); return; }
  let plan;
  try { plan = JSON.parse(raw); }
  catch (e) { showMsg("composeMsg", "invalid JSON: " + e.message, false); return; }
  showMsg("composeMsg", "composing…", true);
  const r = await fetch("/api/compose/run", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ plan }),
  });
  if (!r.ok) {
    const j = await r.json().catch(() => ({}));
    showMsg("composeMsg", (j.error || "compose failed") + (j.stderr ? " — " + j.stderr.slice(0, 300) : ""), false);
    return;
  }
  const blob = await r.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const cd = r.headers.get("Content-Disposition") || "";
  const m = cd.match(/filename="?([^";]+)"?/);
  a.download = m ? m[1] : "deck.pptx";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  showMsg("composeMsg", "downloaded " + a.download, true);
};

// --- brand ---

async function loadBrand() {
  const r = await fetch("/api/compose/brand");
  const j = await r.json();
  document.getElementById("brand").value = j.text || "";
  updateBrandSummary(j.text || "");
}

function updateBrandSummary(text) {
  const trimmed = (text || "").trim();
  const el = document.getElementById("brandSummary");
  if (!trimmed) {
    el.innerHTML = "· <em>empty — not included in bundles</em>";
  } else {
    el.innerHTML = `· <code>${trimmed.length.toLocaleString()} chars</code> active`;
  }
}

document.getElementById("brand").addEventListener("input", (e) => {
  updateBrandSummary(e.target.value);
});

document.getElementById("saveBrand").onclick = async () => {
  const text = document.getElementById("brand").value;
  const r = await fetch("/api/compose/brand", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text }),
  });
  const j = await r.json();
  if (r.ok) {
    showMsg("brandMsg", `saved (${j.chars} chars)`, true);
    updateBrandSummary(text);
  } else {
    showMsg("brandMsg", j.error || "save failed", false);
  }
};

// --- presets ---

async function loadPresets() {
  const r = await fetch("/api/compose/presets");
  const j = await r.json();
  const sel = document.getElementById("presetSelect");
  const cur = sel.value;
  sel.innerHTML = '<option value="">— none —</option>';
  j.presets.forEach((p) => {
    const o = document.createElement("option");
    o.value = p.name;
    o.textContent = p.name;
    if (p.name === cur) o.selected = true;
    sel.appendChild(o);
  });
}

document.getElementById("loadPreset").onclick = async () => {
  const name = document.getElementById("presetSelect").value;
  if (!name) return;
  const r = await fetch("/api/compose/preset?name=" + encodeURIComponent(name));
  if (!r.ok) return;
  const j = await r.json();
  document.getElementById("brief").value = j.text;
};

document.getElementById("savePreset").onclick = async () => {
  const text = document.getElementById("brief").value;
  const name = prompt("Preset name (letters, numbers, _ - space):");
  if (!name) return;
  const r = await fetch("/api/compose/preset", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, text }),
  });
  if (r.ok) {
    await loadPresets();
    document.getElementById("presetSelect").value = name;
  } else {
    const j = await r.json().catch(() => ({}));
    alert(j.error || "save failed");
  }
};

document.getElementById("deletePreset").onclick = async () => {
  const name = document.getElementById("presetSelect").value;
  if (!name) return;
  if (!confirm(`delete preset "${name}"?`)) return;
  await fetch("/api/compose/preset?name=" + encodeURIComponent(name), { method: "DELETE" });
  await loadPresets();
};

(async () => {
  const r = await fetch("/api/compose/options");
  state.options = await r.json();
  renderFilters();
  refreshCount();
  loadBrand();
  loadPresets();
})();
</script>
</body>
</html>
"""


@app.get("/compose")
def compose_page():
    return render_template_string(COMPOSE_HTML)


def main():
    port = int(os.environ.get("PPTX_SKILL_PORT", "5050"))
    url = f"http://127.0.0.1:{port}/"
    Timer(1.2, lambda: webbrowser.open(url)).start()
    print(f"pptx-skill describe app → {url}")
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
