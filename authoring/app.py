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
import time as time_mod
import webbrowser
import zipfile
from collections import deque
from datetime import datetime
from pathlib import Path
from threading import Lock, Timer

import yaml
from flask import Flask, abort, g, jsonify, render_template_string, request, send_file

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
# Debug activity log — in-memory ring buffer surfaced to the UI widget.
# ---------------------------------------------------------------------------

_DEBUG_BUFFER_MAX = 250
_debug_buffer: deque = deque(maxlen=_DEBUG_BUFFER_MAX)
_debug_lock = Lock()
_debug_next_id = 0


def debug_event(level: str, kind: str, msg: str, **details) -> None:
    """Push a named event into the debug ring buffer.

    level: 'info' | 'warn' | 'error'
    kind:  short tag for filtering (e.g. 'http', 'bundle', 'ingest',
           'compose', 'batch', 'user_assets')
    msg:   one-line human-readable summary shown verbatim in the widget
    details: optional structured fields included alongside (not rendered
             by default; available for future drill-downs)
    """
    global _debug_next_id
    with _debug_lock:
        _debug_next_id += 1
        _debug_buffer.append({
            "id": _debug_next_id,
            "ts": time_mod.time(),
            "level": level,
            "kind": kind,
            "msg": msg,
            "details": details or None,
        })


# Skip request logging for our own polling endpoint (would feedback
# loop) and for streamed media we don't care about.
_DEBUG_SKIP_PATHS = ("/api/debug/log",)
_DEBUG_SKIP_PREFIXES = ("/preview", "/api/user_assets/")  # binaries


@app.before_request
def _debug_req_start():
    g._req_started_at = time_mod.time()


@app.after_request
def _debug_req_end(response):
    started = getattr(g, "_req_started_at", None)
    if started is None:
        return response
    path = request.path
    if path in _DEBUG_SKIP_PATHS:
        return response
    if any(path.startswith(p) for p in _DEBUG_SKIP_PREFIXES):
        return response
    # Don't log page renders — the activity is what's interesting.
    if not path.startswith("/api/"):
        return response
    duration_ms = int((time_mod.time() - started) * 1000)
    code = response.status_code
    level = "error" if code >= 500 else ("warn" if code >= 400 else "info")
    debug_event(
        level, "http",
        f"{request.method} {path} → {code} · {duration_ms}ms",
        method=request.method, path=path, status=code,
        duration_ms=duration_ms,
    )
    return response


@app.get("/api/debug/log")
def api_debug_log():
    """Return events newer than `since` (default 0). Cheap snapshot copy
    under the lock; rendering happens client-side."""
    try:
        since = int(request.args.get("since", "0"))
    except ValueError:
        since = 0
    with _debug_lock:
        latest = _debug_next_id
        events = [e for e in _debug_buffer if e["id"] > since]
    return jsonify({"latest_id": latest, "events": events})


@app.post("/api/debug/clear")
def api_debug_clear():
    """Wipe the server-side buffer. Useful when the user wants a clean
    slate before reproducing a specific scenario."""
    global _debug_next_id
    with _debug_lock:
        _debug_buffer.clear()
        # Keep ids monotonic across clears so client polls don't get
        # confused by id wraparound — just bump and continue.
        _debug_next_id += 1
        last = _debug_next_id
    return jsonify({"ok": True, "latest_id": last})


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
            debug_event("error", "ingest",
                        f"ingest failed for {safe_name}: {type(e).__name__}: {e}")
            return jsonify({"error": f"ingest failed: {e}"}), 500
    debug_event(
        "info", "ingest",
        f"ingested {result.get('deck_stem','?')} — "
        f"{result.get('slides',0)} slides, "
        f"{result.get('pictures',0)} pictures, "
        f"{result.get('atoms',0)} atoms",
        **result,
    )
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
    debug_event(
        "info", "batch",
        f"describe batch {batch_id} created — kind={kind}, "
        f"{added}/{count} items, {len(skipped)} skipped"
        + (f", pruned {pruned} old" if pruned else ""),
        batch_id=batch_id, kind=kind, count=added, requested=count,
        skipped=len(skipped),
    )

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
    matched = sum(1 for r in results if r["status"] not in ("no-match",))
    n_err = sum(1 for r in results if r.get("errors"))
    debug_event(
        "warn" if n_err else "info", "batch",
        f"batch {batch_id} applied — {matched} matched, {n_err} with errors",
        batch_id=batch_id, matched=matched, errors=n_err,
    )
    return jsonify({
        "batch_id": batch_id,
        "results": results,
        "found_keys": found_keys,
        "matched": matched,
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
        + "# How to plan this deck\n\n"
        "Work through the problem in **four passes**, in this order. The\n"
        "Compose page extracts ONLY the final JSON code-fence at the end\n"
        "of your response; everything else is reasoning prose that helps\n"
        "you think clearly and helps the human spot mistakes. Show your\n"
        "work — don't jump straight to JSON.\n"
        "\n"
        "## Pass 1 — outline (plain text)\n"
        "\n"
        "Before picking any templates, sketch the deck as a numbered list\n"
        "of slide titles + one-line intents. Decide the narrative arc and\n"
        "slide count from the brief. Output under the heading `## Outline`.\n"
        "\n"
        "Example shape:\n"
        "```\n"
        "## Outline\n"
        "1. Opener — thesis title + author + university\n"
        "2. Problem statement — why this matters\n"
        "3. Methodology — how the study was conducted\n"
        "4. Results — key findings + chart\n"
        "5. Closing — implications + Q&A invitation\n"
        "```\n"
        "\n"
        "## Pass 2 — template picks\n"
        "\n"
        "For each outline entry, pick ONE template id from `index.json`.\n"
        "Justify each pick in one line: which `feel` / `suitable_for` /\n"
        "inventory anatomy matched, and how the template's `theme_colors`\n"
        "compare to brand policy (see `brand.md` if present).\n"
        "\n"
        "Use `helpers/kb_filter.py` to find candidates, then\n"
        "`helpers/kb_inspect.py <template_id>` to read the picked\n"
        "template's slots + inventory in denormalized form. Pick templates\n"
        "ONLY from `index.json` (or from `user_assets/manifest.json` for\n"
        "assets — see the user-assets section above when present).\n"
        "\n"
        "If a brief calls for a layout no existing template provides,\n"
        "switch that entry to compose-mode\n"
        "(`{\"compose\": true, \"layout\": \"...\", \"shapes\": [...]}`)\n"
        "rather than forcing a template that doesn't fit.\n"
        "\n"
        "Output as a table or list under `## Picks`.\n"
        "\n"
        "## Pass 3 — slot values\n"
        "\n"
        "Now fill each pick's slots. Walk `helpers/kb_inspect.py\n"
        "<template_id>` to see what slots the template exposes + their\n"
        "constraints, then:\n"
        "\n"
        "- **Text slots** — plain string, respect each slot's `max_chars`.\n"
        "  Use `helpers/kb_budget.py <template_id> <slot_id> \"draft\"` to\n"
        "  check fit on tight slots.\n"
        "- **Bullets slots** — array of plain strings. NO leading `•`, `-`,\n"
        "  or `*` glyphs — the template applies bullets via layout.\n"
        "- **Image slots** — an `asset_<id>` from `index.json`, or from\n"
        "  `user_assets/manifest.json` when present. User-supplied assets\n"
        "  outrank KB matches.\n"
        "- If no asset fits a slot, omit the slot rather than forcing one.\n"
        "\n"
        "## Pass 4 — self-lint\n"
        "\n"
        "Save your draft as `plan.json` and run\n"
        "`python helpers/kb_lint.py < plan.json` (or pipe directly). The\n"
        "linter catches: slot ids not declared on the template, text over\n"
        "`max_chars`, leading bullet glyphs, asset ids missing from both\n"
        "index.json and user_assets, accepted-but-degraded shapes, and\n"
        "compose-mode entries that the engine will currently skip. Fix\n"
        "errors before emitting the final JSON.\n"
        "\n"
        "Exit code 0 = clean, 1 = errors — must be 0 to consider the plan\n"
        "ready.\n"
        "\n"
        "## Final output\n"
        "\n"
        "End your response with the plan inside a single fenced block,\n"
        "exactly like this (the fence is what the Compose page extracts):\n"
        "\n"
        "    ```json\n"
        "    [\n"
        "      {\"template\": \"<id>\", \"slots\": { ... }},\n"
        "      ...\n"
        "    ]\n"
        "    ```\n"
        "\n"
        "Any prose before the fence is ignored by the system; any prose\n"
        "after it is also ignored. Put only one JSON fence per response.\n"
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

    # User-supplied assets from the last bundle generation: copy the
    # full-res originals into staging/assets/ alongside KB assets so
    # reader.py resolves them transparently. Write a minimal yaml stub
    # (kind + subject) — compose-mode needs the sidecar to exist; image
    # slots don't read it but the file being present is harmless.
    user_meta = _read_user_meta(USER_BUNDLE_DIR)
    for aid, entry in user_meta.items():
        ext = entry.get("ext") or "bin"
        src = USER_BUNDLE_DIR / f"{aid}.{ext}"
        if not src.exists():
            continue
        # Don't overwrite a KB asset with the same id (extremely unlikely
        # — would mean SHA1 collision on file content — but be defensive).
        dst = ast_dir / f"{aid}.{ext}"
        if dst.exists():
            continue
        shutil.copyfile(src, dst)
        stub = {
            "id": aid,
            "kind": entry.get("kind") or "image",
            "subject": f"User-supplied {entry.get('kind') or 'asset'} "
                       f"({entry.get('filename') or '?'})",
            "sources": [],
        }
        (ast_dir / f"{aid}.yaml").write_text(
            yaml.safe_dump(stub, sort_keys=False, allow_unicode=True),
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
        debug_event("warn", "bundle",
                    "bundle request rejected — filters match nothing")
        return jsonify({"error": "filters match nothing — broaden them"}), 400
    user_count = len(_read_user_meta(USER_STAGED_DIR))
    blob = _build_prompt_bundle_zip(slides, assets, brief)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    debug_event(
        "info", "bundle",
        f"prompt bundle built — {len(slides)} templates, "
        f"{len(assets)} assets, {user_count} user assets, "
        f"{len(blob) // 1024} KB",
        templates=len(slides), assets=len(assets),
        user_assets=user_count, size_bytes=len(blob),
    )
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
    if added:
        names = ", ".join(e.get("filename", "?") for e in added)
        debug_event(
            "info", "user_assets",
            f"uploaded {len(added)} file(s): {names}",
            count=len(added),
        )
    if errors:
        debug_event(
            "warn", "user_assets",
            f"rejected {len(errors)} upload(s)",
            errors=errors,
        )
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
    debug_event(
        "info", "user_assets",
        f"removed staged user asset {asset_id} ({entry.get('filename','?')})",
    )
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


def _plan_looks_like_v5(plan: list) -> bool:
    """A v5 plan entry has `skeleton_id`. v4 entries use `template` or
    `compose: true`. Detect by the *first* entry that's a dict — keeps
    detection cheap and consistent across the whole plan."""
    for entry in plan:
        if isinstance(entry, dict):
            return "skeleton_id" in entry
    return False


def _stage_compose_v5_bundle(staging: Path) -> None:
    """Stage a v5 skill bundle under `staging` for compose-v5 to read.

    Layout mirrors cli.py build-v5 (themes/, skeletons/, assets/,
    reader.py, index.json) so the subprocess sees the same shape it'd
    get from an unzipped skill-v5.zip. SKILL.md is omitted — compose-
    v5 doesn't read it, and shipping it would just bloat the staging
    dir.
    """
    consumer_reader = cli_mod.CONSUMER / "reader.py"
    (staging / "reader.py").write_text(
        consumer_reader.read_text(encoding="utf-8"), encoding="utf-8",
    )

    themes_root = cli_mod.WORKSPACE / "themes"
    skeletons_root = cli_mod.WORKSPACE / "skeletons"
    assets_root = cli_mod.WORKSPACE / "assets"

    # Themes
    themes_summary = []
    if themes_root.exists():
        for d in sorted(themes_root.iterdir()):
            if not d.is_dir() or not (d / "theme.yaml").exists():
                continue
            t = cli_mod.read_yaml(d / "theme.yaml")
            tid = t.get("id", d.name)
            t_dir = staging / "themes" / tid
            t_dir.mkdir(parents=True, exist_ok=True)
            for fn in ("theme.yaml", "master.pptx", "preview.png"):
                src = d / fn
                if src.exists():
                    shutil.copyfile(src, t_dir / fn)
            themes_summary.append({"id": tid, "palette": t.get("palette", {}),
                                   "fonts": t.get("fonts", {})})

    # Skeletons (skip rejected ones — same filter as build-v5)
    skeletons_summary = []
    if skeletons_root.exists():
        for d in sorted(skeletons_root.iterdir()):
            if not d.is_dir() or not (d / "skeleton.yaml").exists():
                continue
            sk = cli_mod.read_yaml(d / "skeleton.yaml")
            if sk.get("status") == "rejected":
                continue
            sid = sk.get("id", d.name)
            sk_dir = staging / "skeletons" / sid
            sk_dir.mkdir(parents=True, exist_ok=True)
            for fn in ("skeleton.yaml", "preview.png", "background.png"):
                src = d / fn
                if src.exists():
                    shutil.copyfile(src, sk_dir / fn)
            skeletons_summary.append({
                "id": sid,
                "source_deck": sk.get("source_deck"),
                "categories": sk.get("categories") or [],
                "slot_count": len(sk.get("slots") or []),
                "slot_kinds": sorted({s.get("kind") for s in (sk.get("slots") or [])}),
                "status": sk.get("status", "pending"),
            })

    # Assets (binary + sidecar). Re-use the v4 staging logic so user-
    # supplied assets ride along too.
    ast_dir = staging / "assets"
    ast_dir.mkdir(parents=True, exist_ok=True)
    asset_summary = []
    if assets_root.exists():
        for yaml_path in sorted(assets_root.glob("*.yaml")):
            ad = cli_mod.read_yaml(yaml_path)
            aid = ad.get("id")
            if not aid:
                continue
            binary = _asset_binary(yaml_path)
            if binary is not None:
                ext = binary.suffix.lstrip(".") or "bin"
                shutil.copyfile(binary, ast_dir / f"{aid}.{ext}")
            clean = {k: v for k, v in ad.items() if not k.startswith("_")}
            (ast_dir / f"{aid}.yaml").write_text(
                yaml.safe_dump(clean, sort_keys=False, allow_unicode=True),
                encoding="utf-8",
            )
            asset_summary.append({"id": aid, "kind": ad.get("kind"),
                                  "subject": ad.get("subject", "")})

    # User-supplied assets from the last bundle generation — mirror
    # _stage_compose_bundle so v5 plans referencing user assets work too.
    user_meta = _read_user_meta(USER_BUNDLE_DIR)
    for aid, entry in user_meta.items():
        ext = entry.get("ext") or "bin"
        src = USER_BUNDLE_DIR / f"{aid}.{ext}"
        if not src.exists():
            continue
        dst = ast_dir / f"{aid}.{ext}"
        if dst.exists():
            continue
        shutil.copyfile(src, dst)
        stub = {
            "id": aid,
            "kind": entry.get("kind") or "image",
            "subject": f"User-supplied {entry.get('kind') or 'asset'} "
                       f"({entry.get('filename') or '?'})",
            "sources": [],
        }
        (ast_dir / f"{aid}.yaml").write_text(
            yaml.safe_dump(stub, sort_keys=False, allow_unicode=True),
            encoding="utf-8",
        )

    # index.json — v5 shape, mirroring build-v5
    index = {
        "version": 5,
        "themes": themes_summary,
        "skeletons": skeletons_summary,
        "assets": asset_summary,
    }
    (staging / "index.json").write_text(
        json_mod.dumps(index, indent=2, ensure_ascii=False), encoding="utf-8",
    )


def _pick_v5_theme(plan: list, staging: Path) -> tuple[str | None, str]:
    """Pick a host theme for a v5 plan.

    Strategy:
      1. If any plan entry includes an explicit "theme" key, use that.
      2. Else look up the first plan entry's skeleton, take its
         source_deck — if a theme with that id exists, use it.
      3. Else fall back to the first theme alphabetically.

    Returns (theme_id, reason). theme_id is None when no theme is
    available — caller should error with reason.
    """
    themes_dir = staging / "themes"
    available = sorted([d.name for d in themes_dir.iterdir() if d.is_dir()]) \
        if themes_dir.exists() else []
    if not available:
        return None, "no themes in workspace — run ingest first"

    # 1. Explicit theme in plan
    for entry in plan:
        if isinstance(entry, dict) and entry.get("theme"):
            t = entry["theme"]
            if t in available:
                return t, f"explicit theme {t!r} from plan"
            return None, f"plan requested theme {t!r} but only {available} available"

    # 2. Match first skeleton's source_deck
    first_sk_id = next((e.get("skeleton_id") for e in plan
                        if isinstance(e, dict) and e.get("skeleton_id")), None)
    if first_sk_id:
        sk_yaml = staging / "skeletons" / first_sk_id / "skeleton.yaml"
        if sk_yaml.exists():
            try:
                sk = cli_mod.read_yaml(sk_yaml)
                deck = sk.get("source_deck")
                if deck and deck in available:
                    return deck, f"matched source_deck {deck!r} of skeleton {first_sk_id!r}"
            except Exception:
                pass

    # 3. First alphabetically
    return available[0], f"defaulted to first theme {available[0]!r}"


@app.post("/api/compose/run")
def api_compose_run():
    body = request.get_json(force=True) or {}
    plan = body.get("plan")
    if not isinstance(plan, list) or not plan:
        return jsonify({"error": "plan must be a non-empty JSON array"}), 400

    is_v5 = _plan_looks_like_v5(plan)

    with tempfile.TemporaryDirectory(prefix="pptx-compose-") as tmpdir:
        staging = Path(tmpdir) / "bundle"
        staging.mkdir(parents=True, exist_ok=True)

        if is_v5:
            _stage_compose_v5_bundle(staging)
            theme_id, theme_reason = _pick_v5_theme(plan, staging)
            if theme_id is None:
                return jsonify({"error": f"v5 compose blocked: {theme_reason}"}), 400
            # Strip the optional "theme" key so reader's plan validator
            # (which doesn't know about envelope fields) stays happy.
            cleaned = [{k: v for k, v in e.items() if k != "theme"}
                       if isinstance(e, dict) else e for e in plan]
            plan_path = staging / "plan.json"
            plan_path.write_text(
                json_mod.dumps(cleaned, ensure_ascii=False), encoding="utf-8",
            )
            cmd = [sys.executable, "reader.py", "compose-v5",
                   "plan.json", "out.pptx", "--theme", theme_id]
            debug_event("info", "compose",
                        f"v5 compose: {theme_reason}, {len(plan)} slide(s)")
        else:
            _stage_compose_bundle(staging)
            plan_path = staging / "plan.json"
            plan_path.write_text(
                json_mod.dumps(plan, ensure_ascii=False), encoding="utf-8",
            )
            cmd = [sys.executable, "reader.py", "compose", "plan.json", "out.pptx"]

        out_path = staging / "out.pptx"
        try:
            result = subprocess.run(
                cmd, cwd=str(staging),
                capture_output=True, text=True, timeout=60,
            )
        except subprocess.TimeoutExpired:
            debug_event("error", "compose", "compose subprocess timed out (60s)")
            return jsonify({"error": "compose timed out"}), 500
        if result.returncode != 0 or not out_path.exists():
            stderr_head = (result.stderr or "").splitlines()
            debug_event(
                "error", "compose",
                f"compose failed (exit {result.returncode}): "
                + (stderr_head[-1] if stderr_head else "no stderr"),
            )
            return jsonify({
                "error": "compose failed",
                "stdout": result.stdout,
                "stderr": result.stderr,
                "mode": "v5" if is_v5 else "v4",
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
        suffix = "v5" if is_v5 else "v4"
        out_persisted = persisted / f"deck-{suffix}-{ts}.pptx"
        shutil.copyfile(out_path, out_persisted)
        debug_event(
            "info", "compose",
            f"compose finished ({suffix}) — {out_persisted.name}, "
            f"{out_persisted.stat().st_size // 1024} KB, "
            f"{len(plan)} plan entries",
            file=out_persisted.name, size_bytes=out_persisted.stat().st_size,
            entries=len(plan), mode=suffix,
        )

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


DEBUG_WIDGET = r"""
<style>
.dbg-floater { position: fixed; bottom: 14px; right: 14px; z-index: 99999;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
.dbg-toggle { padding: 7px 13px; background: #222; color: white; border: none;
               border-radius: 999px; cursor: pointer; font-size: 12px;
               font-weight: 600; box-shadow: 0 2px 6px rgba(0,0,0,0.2);
               display: flex; align-items: center; gap: 6px; }
.dbg-toggle:hover { background: #000; }
.dbg-badge { background: #0066cc; color: white; border-radius: 10px;
              padding: 1px 7px; font-size: 11px; font-weight: 600;
              min-width: 14px; text-align: center; }
.dbg-badge.err { background: #c92a2a; }
.dbg-panel { position: absolute; bottom: 44px; right: 0; width: 480px;
              height: 320px; background: white; border: 1px solid #ddd;
              border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.18);
              display: flex; flex-direction: column; overflow: hidden; }
.dbg-header { padding: 7px 12px; border-bottom: 1px solid #e5e5e5;
               background: #f6f7f8; display: flex; align-items: center;
               font-size: 12px; gap: 6px; }
.dbg-header strong { flex: 1; color: #333; }
.dbg-header button { background: none; border: none; cursor: pointer;
                      color: #666; font-size: 12px; padding: 2px 8px;
                      border-radius: 3px; }
.dbg-header button:hover { background: #ececec; color: #222; }
.dbg-list { list-style: none; padding: 0; margin: 0; flex: 1;
             overflow-y: auto; font-size: 11px;
             font-family: ui-monospace, Menlo, monospace; }
.dbg-row { padding: 3px 12px; border-bottom: 1px solid #f4f4f4;
            display: flex; gap: 8px; align-items: baseline; }
.dbg-row.dbg-warn { background: #fff8e6; }
.dbg-row.dbg-error { background: #fdecea; color: #8a1818; }
.dbg-ts { color: #aaa; flex-shrink: 0; }
.dbg-kind { color: #0066cc; font-weight: 600; flex-shrink: 0;
             min-width: 70px; }
.dbg-row.dbg-error .dbg-kind { color: #8a1818; }
.dbg-msg { color: #333; word-break: break-word; flex: 1; }
.dbg-row.dbg-error .dbg-msg { color: #8a1818; }
.dbg-empty { padding: 20px; text-align: center; color: #aaa; font-size: 11px; }
</style>
<div class="dbg-floater" id="dbgRoot">
  <button class="dbg-toggle" id="dbgToggle" type="button" title="Toggle activity log">
    <span>Activity</span>
    <span class="dbg-badge" id="dbgBadge" hidden></span>
  </button>
  <div class="dbg-panel" id="dbgPanel" hidden>
    <div class="dbg-header">
      <strong>Activity log</strong>
      <button id="dbgClear" type="button" title="Clear">clear</button>
      <button id="dbgClose" type="button" title="Hide">×</button>
    </div>
    <ul class="dbg-list" id="dbgList">
      <li class="dbg-empty" id="dbgEmpty">(waiting for activity…)</li>
    </ul>
  </div>
</div>
<script>
(() => {
  const root = document.getElementById("dbgRoot");
  if (!root) return;
  const list = document.getElementById("dbgList");
  const panel = document.getElementById("dbgPanel");
  const toggle = document.getElementById("dbgToggle");
  const badge = document.getElementById("dbgBadge");
  const emptyEl = document.getElementById("dbgEmpty");
  let lastId = 0;
  let panelOpen = false;
  let unread = 0;
  let unreadErr = 0;

  function fmtTime(ts) {
    const d = new Date(ts * 1000);
    return d.toLocaleTimeString();
  }
  function updateBadge() {
    if (unread > 0) {
      badge.textContent = unreadErr > 0 ? (unread + "!") : String(unread);
      badge.className = "dbg-badge" + (unreadErr > 0 ? " err" : "");
      badge.hidden = false;
    } else {
      badge.hidden = true;
    }
  }
  function appendEvents(events) {
    if (!events.length) return;
    if (emptyEl && emptyEl.parentNode) emptyEl.remove();
    const wasNearBottom =
      panel.scrollHeight - panel.scrollTop - panel.clientHeight < 80;
    for (const e of events) {
      const li = document.createElement("li");
      li.className = "dbg-row dbg-" + (e.level || "info");
      const ts = document.createElement("span");
      ts.className = "dbg-ts";
      ts.textContent = fmtTime(e.ts);
      const kind = document.createElement("span");
      kind.className = "dbg-kind";
      kind.textContent = e.kind || "?";
      const msg = document.createElement("span");
      msg.className = "dbg-msg";
      msg.textContent = e.msg || "";
      li.append(ts, kind, msg);
      list.appendChild(li);
    }
    // Trim if grown too long.
    while (list.children.length > 400) list.removeChild(list.firstChild);
    if (panelOpen && wasNearBottom) {
      const inner = panel.querySelector(".dbg-list");
      if (inner) inner.scrollTop = inner.scrollHeight;
    }
  }
  async function poll() {
    try {
      const r = await fetch("/api/debug/log?since=" + lastId);
      if (!r.ok) return;
      const j = await r.json();
      if (j.latest_id > lastId) lastId = j.latest_id;
      if (j.events && j.events.length) {
        appendEvents(j.events);
        if (!panelOpen) {
          for (const e of j.events) {
            unread += 1;
            if (e.level === "error") unreadErr += 1;
          }
          updateBadge();
        }
      }
    } catch (err) {
      // Silent — backend may be down briefly.
    }
  }
  toggle.onclick = () => {
    if (panelOpen) {
      panel.hidden = true; panelOpen = false;
    } else {
      panel.hidden = false; panelOpen = true;
      unread = 0; unreadErr = 0;
      updateBadge();
      const inner = panel.querySelector(".dbg-list");
      if (inner) inner.scrollTop = inner.scrollHeight;
    }
  };
  document.getElementById("dbgClose").onclick = () => {
    panel.hidden = true; panelOpen = false;
  };
  document.getElementById("dbgClear").onclick = async () => {
    list.innerHTML = '<li class="dbg-empty" id="dbgEmpty">(cleared)</li>';
    unread = 0; unreadErr = 0;
    updateBadge();
    try {
      const r = await fetch("/api/debug/clear", { method: "POST" });
      if (r.ok) {
        const j = await r.json();
        if (j.latest_id) lastId = j.latest_id;
      }
    } catch (err) {}
  };
  poll();
  setInterval(poll, 2000);
})();
</script>
"""


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
      <span style="float:right;font-size:11px;">
        <a href="/v5" style="color:#0066cc;text-decoration:none;margin-right:8px;">Skeletons →</a>
        <a href="/compose" style="color:#0066cc;text-decoration:none;">Compose →</a>
      </span>
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
      <button data-tab="slides" title="Slides moved to /v5 — describe page now owns assets only"
              style="display:none;">Slides</button>
      <button data-tab="assets" class="active">Assets</button>
    </div>
    <div style="padding:6px 14px;font-size:11px;color:#888;border-bottom:1px solid #eee;
                line-height:1.4;">
      Slide review moved to <a href="/v5" style="color:#0066cc;text-decoration:none;">/v5 →</a>
      — this page now describes assets only.
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

let activeTab = "assets";  // v5: slide tab hidden; describe page owns assets only
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
{{ debug_widget|safe }}
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(INDEX_HTML, debug_widget="")


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
    .ua-block { margin-top: 12px; }
    .ua-header { display: flex; align-items: center; gap: 10px;
                  font-size: 13px; color: #333; margin-bottom: 6px; }
    .ua-header .ua-label { font-weight: 600; }
    .ua-header .ua-counter { color: #888; font-size: 12px; margin-left: auto; }
    .ua-dropzone { border: 1.5px dashed #c5d4ea; border-radius: 6px;
                    padding: 14px; text-align: center; background: #fafcff;
                    color: #555; font-size: 12px; cursor: pointer;
                    transition: background 0.15s, border-color 0.15s; }
    .ua-dropzone:hover, .ua-dropzone.drag-over {
        background: #eef4ff; border-color: #0066cc; }
    .ua-dropzone strong { color: #0066cc; }
    .ua-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(120px, 1fr));
                gap: 10px; margin-top: 10px; }
    .ua-tile { border: 1px solid #ddd; border-radius: 6px; overflow: hidden;
                background: white; display: flex; flex-direction: column;
                position: relative; }
    .ua-tile .thumb { height: 80px; background: #f5f5f5; display: flex;
                       align-items: center; justify-content: center;
                       overflow: hidden; }
    .ua-tile .thumb img { max-width: 100%; max-height: 100%; object-fit: contain; }
    .ua-tile .thumb .non-img { font-size: 11px; color: #888;
                                 font-family: ui-monospace, monospace; }
    .ua-tile .meta { padding: 6px 8px; font-size: 11px; }
    .ua-tile .meta .fn { font-weight: 600; color: #333; white-space: nowrap;
                          overflow: hidden; text-overflow: ellipsis; }
    .ua-tile .meta .sub { color: #888; margin-top: 2px; }
    .ua-tile .remove { position: absolute; top: 4px; right: 4px;
                        width: 22px; height: 22px; border-radius: 50%;
                        background: rgba(0,0,0,0.55); color: white;
                        border: none; cursor: pointer; font-size: 13px;
                        line-height: 1; display: flex; align-items: center;
                        justify-content: center; padding: 0; }
    .ua-tile .remove:hover { background: rgba(180,0,0,0.85); }
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
    <a href="/v5">v5 skeletons →</a>
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

      <div class="ua-block">
        <div class="ua-header">
          <span class="ua-label">Attach your own assets (optional)</span>
          <span class="ua-counter" id="uaCounter">0 files</span>
        </div>
        <div class="ua-dropzone" id="uaDropzone">
          <strong>Click to choose files</strong> or drop here —
          png / jpg / webp / gif / svg / xml.<br>
          Sent to the agent as low-res previews; originals stay on this
          machine and are spliced into the deck at compose time.
        </div>
        <input type="file" id="uaFileInput" multiple
               accept=".png,.jpg,.jpeg,.webp,.gif,.svg,.xml"
               style="display:none">
        <div class="ua-grid" id="uaGrid"></div>
      </div>

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
        Paste the JSON array the LLM returns. v4 plans
        (<code>{"template": …, "slots": …}</code>) run via
        <code>reader.py compose</code>. v5 plans
        (<code>{"skeleton_id": …, "slots": …}</code>) auto-route to
        <code>reader.py compose-v5</code>; host theme is picked from
        the first skeleton's <code>source_deck</code> unless a plan
        entry sets <code>"theme": "&lt;id&gt;"</code>.
      </div>
      <textarea id="plan" class="plan-area" placeholder='[{"skeleton_id":"deckA_03","slots":{"title":"…"}}]'></textarea>
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
  // Staged user assets just moved into the bundle snapshot — refresh the
  // grid so the user sees the staged area cleared (their files are now
  // persisted next to the saved bundle for the compose round-trip).
  loadUserAssets();
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

// Extract a JSON array from arbitrary pasted text. Tolerates:
//   - bare JSON: `[ {...}, {...} ]`
//   - fenced JSON: `... reasoning prose ... ```json\n[ ... ]\n``` ... trailing prose`
//   - reasoning + bare array on the last lines.
// Strategy: try straight parse first; then a ```json fence; then a `json
// fenced block without language tag; finally fall back to the substring
// from the first `[` to the matching last `]`. Returns the parsed array
// or throws with a descriptive message.
function extractPlanJSON(raw) {
  // 1. Straight parse.
  try { return JSON.parse(raw); } catch (e) {}

  // 2. ```json ... ``` (case-insensitive language tag).
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenced) {
    try { return JSON.parse(fenced[1].trim()); } catch (e) {}
  }

  // 3. First `[` to last `]` substring.
  const first = raw.indexOf("[");
  const last = raw.lastIndexOf("]");
  if (first !== -1 && last > first) {
    try { return JSON.parse(raw.slice(first, last + 1)); } catch (e) {}
  }

  throw new Error(
    "could not extract a JSON array — paste the model's response " +
    "verbatim (with or without a ```json fence), or just the bare array."
  );
}

document.getElementById("runCompose").onclick = async () => {
  const raw = document.getElementById("plan").value.trim();
  if (!raw) { showMsg("composeMsg", "paste a plan first", false); return; }
  let plan;
  try { plan = extractPlanJSON(raw); }
  catch (e) { showMsg("composeMsg", e.message, false); return; }
  if (!Array.isArray(plan)) {
    showMsg("composeMsg", "extracted JSON is not an array", false); return;
  }
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

// --- User-supplied assets ----------------------------------------------

function _fmtBytes(n) {
  if (!n) return "0 B";
  if (n < 1024) return n + " B";
  if (n < 1024 * 1024) return Math.round(n / 1024) + " KB";
  return (n / 1024 / 1024).toFixed(1) + " MB";
}

function renderUserAssets(staged) {
  const grid = document.getElementById("uaGrid");
  grid.innerHTML = "";
  const counter = document.getElementById("uaCounter");
  const totalBytes = staged.reduce((a, e) => a + (e.size_bytes || 0), 0);
  counter.textContent = `${staged.length} file${staged.length === 1 ? "" : "s"} · ${_fmtBytes(totalBytes)}`;
  for (const e of staged) {
    const tile = document.createElement("div");
    tile.className = "ua-tile";
    const thumb = document.createElement("div");
    thumb.className = "thumb";
    const previewable = (e.kind === "image" || e.kind === "vector");
    if (previewable) {
      const img = document.createElement("img");
      img.src = `/api/user_assets/${e.id}/preview`;
      img.alt = e.filename || e.id;
      thumb.appendChild(img);
    } else {
      const placeholder = document.createElement("div");
      placeholder.className = "non-img";
      placeholder.textContent = (e.ext || "?").toUpperCase();
      thumb.appendChild(placeholder);
    }
    tile.appendChild(thumb);
    const meta = document.createElement("div");
    meta.className = "meta";
    const fn = document.createElement("div");
    fn.className = "fn";
    fn.textContent = e.filename || e.id;
    fn.title = e.filename || e.id;
    meta.appendChild(fn);
    const sub = document.createElement("div");
    sub.className = "sub";
    const dims = (e.width && e.height) ? `${e.width}×${e.height}px · ` : "";
    sub.textContent = `${dims}${_fmtBytes(e.size_bytes)}`;
    meta.appendChild(sub);
    tile.appendChild(meta);
    const rm = document.createElement("button");
    rm.className = "remove";
    rm.title = "Remove";
    rm.textContent = "×";
    rm.onclick = async () => {
      rm.disabled = true;
      await fetch(`/api/user_assets/${e.id}`, { method: "DELETE" });
      await loadUserAssets();
    };
    tile.appendChild(rm);
    grid.appendChild(tile);
  }
}

async function loadUserAssets() {
  const r = await fetch("/api/user_assets");
  if (!r.ok) { renderUserAssets([]); return; }
  const j = await r.json();
  renderUserAssets(j.staged || []);
}

async function uploadUserAssets(files) {
  if (!files || !files.length) return;
  const fd = new FormData();
  for (const f of files) fd.append("files", f);
  showMsg("bundleMsg", `uploading ${files.length} file(s)…`, true);
  const r = await fetch("/api/user_assets", { method: "POST", body: fd });
  if (!r.ok) {
    showMsg("bundleMsg", "upload failed", false);
    return;
  }
  const j = await r.json();
  await loadUserAssets();
  if (j.errors && j.errors.length) {
    const reasons = j.errors.map(e => `${e.filename}: ${e.reason}`).join("; ");
    showMsg("bundleMsg", `added ${j.added.length}, rejected: ${reasons}`, false);
  } else {
    showMsg("bundleMsg", `added ${j.added.length} file(s)`, true);
  }
}

document.getElementById("uaDropzone").onclick = () => {
  document.getElementById("uaFileInput").click();
};
document.getElementById("uaFileInput").onchange = (e) => {
  uploadUserAssets(e.target.files);
  e.target.value = "";  // allow re-selecting the same file later
};
(() => {
  const dz = document.getElementById("uaDropzone");
  ["dragenter", "dragover"].forEach(ev =>
    dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add("drag-over"); })
  );
  ["dragleave", "drop"].forEach(ev =>
    dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove("drag-over"); })
  );
  dz.addEventListener("drop", e => {
    if (e.dataTransfer && e.dataTransfer.files) {
      uploadUserAssets(e.dataTransfer.files);
    }
  });
})();

(async () => {
  const r = await fetch("/api/compose/options");
  state.options = await r.json();
  renderFilters();
  refreshCount();
  loadBrand();
  loadPresets();
  loadUserAssets();
})();
</script>
{{ debug_widget|safe }}
</body>
</html>
"""


@app.get("/compose")
def compose_page():
    return render_template_string(COMPOSE_HTML, debug_widget="")


# ---------------------------------------------------------------------------
# v5 redesign — read-only skeletons view (phase C1).
# Self-contained block. To roll back v5, delete this section + the
# V5_HTML constant + the route below.
# ---------------------------------------------------------------------------


V5_THEMES_DIR = WORKSPACE / "themes"
V5_SKELETONS_DIR = WORKSPACE / "skeletons"


def _v5_safe_path(root: Path, name: str) -> Path | None:
    """Resolve <root>/<name> ensuring no traversal outside <root>."""
    if not name or "/" in name or "\\" in name or ".." in name:
        return None
    target = (root / name).resolve()
    try:
        target.relative_to(root.resolve())
    except ValueError:
        return None
    return target


@app.get("/api/v5/skeletons")
def api_v5_list_skeletons():
    """Group skeleton summaries by source_deck."""
    decks: dict = {}
    if not V5_SKELETONS_DIR.exists():
        return jsonify({"decks": decks})
    for d in sorted(V5_SKELETONS_DIR.iterdir()):
        if not d.is_dir():
            continue
        sk_path = d / "skeleton.yaml"
        if not sk_path.exists():
            continue
        try:
            data = yaml.safe_load(sk_path.read_text(encoding="utf-8")) or {}
        except Exception:
            continue
        deck = data.get("source_deck", "unknown")
        decks.setdefault(deck, []).append({
            "id": data.get("id"),
            "source_slide_index": data.get("source_slide_index"),
            "status": data.get("status", "pending"),
            "categories": data.get("categories", []),
            "has_warnings": bool(data.get("digest_warnings")),
            "slot_count": len(data.get("slots", [])),
            "has_preview": (d / "preview.png").exists(),
        })
    for deck in decks:
        decks[deck].sort(key=lambda s: s.get("source_slide_index", 0))
    return jsonify({"decks": decks})


@app.get("/api/v5/skeleton/<skeleton_id>")
def api_v5_get_skeleton(skeleton_id):
    safe = _v5_safe_path(V5_SKELETONS_DIR, skeleton_id)
    if safe is None or not (safe / "skeleton.yaml").exists():
        abort(404)
    try:
        data = yaml.safe_load((safe / "skeleton.yaml").read_text(encoding="utf-8")) or {}
    except Exception:
        abort(500)
    return jsonify(data)


@app.get("/v5/skeleton/<skeleton_id>/preview.png")
def v5_skeleton_preview(skeleton_id):
    safe = _v5_safe_path(V5_SKELETONS_DIR, skeleton_id)
    if safe is None:
        abort(404)
    p = safe / "preview.png"
    if not p.exists():
        abort(404)
    return send_file(str(p), mimetype="image/png")


@app.get("/api/v5/themes")
def api_v5_list_themes():
    out: list[dict] = []
    if not V5_THEMES_DIR.exists():
        return jsonify({"themes": out})
    for d in sorted(V5_THEMES_DIR.iterdir()):
        if not d.is_dir():
            continue
        t_path = d / "theme.yaml"
        if not t_path.exists():
            continue
        try:
            data = yaml.safe_load(t_path.read_text(encoding="utf-8")) or {}
        except Exception:
            continue
        out.append({
            "id": data.get("id"),
            "palette": data.get("palette", {}),
            "fonts": data.get("fonts", {}),
            "decoration_count": len(data.get("decorations", [])),
        })
    return jsonify({"themes": out})


@app.get("/api/v5/theme/<theme_id>")
def api_v5_get_theme(theme_id):
    safe = _v5_safe_path(V5_THEMES_DIR, theme_id)
    if safe is None or not (safe / "theme.yaml").exists():
        abort(404)
    try:
        data = yaml.safe_load((safe / "theme.yaml").read_text(encoding="utf-8")) or {}
    except Exception:
        abort(500)
    return jsonify(data)


@app.get("/v5")
def v5_page():
    # Debug widget intentionally omitted — it's a v4 compose-flow tool
    # and clutters the skeleton review with no upside on this page.
    return render_template_string(V5_HTML)


# --- C-actions: write-back endpoints ------------------------------------


_V5_VALID_STATUSES = {"pending", "done", "rejected"}
_V5_VALID_OVERLAP_DECISIONS = {"image_slot", "reject", "freeze_pending"}
_V5_VALID_KINDS = {"heading", "paragraph", "bullets", "image", "table", "chart", "footer"}


def _v5_load_skeleton(skeleton_id: str) -> tuple[Path | None, dict | None, str | None]:
    """Resolve <id> safely and load skeleton.yaml. Returns (yaml_path,
    data, error_message). On any failure path returns (None, None, msg)
    suitable for the caller to JSON-error.
    """
    safe = _v5_safe_path(V5_SKELETONS_DIR, skeleton_id)
    if safe is None:
        return None, None, "invalid skeleton id"
    p = safe / "skeleton.yaml"
    if not p.exists():
        return None, None, "skeleton not found"
    try:
        return p, yaml.safe_load(p.read_text(encoding="utf-8")) or {}, None
    except Exception as e:
        return None, None, f"failed to read skeleton.yaml: {e}"


def _v5_save_skeleton(p: Path, data: dict) -> None:
    p.write_text(
        yaml.safe_dump(data, sort_keys=False, allow_unicode=True, width=100),
        encoding="utf-8",
    )


@app.post("/api/v5/skeleton/<skeleton_id>/status")
def api_v5_set_status(skeleton_id):
    p, data, err = _v5_load_skeleton(skeleton_id)
    if err:
        return jsonify({"error": err}), 404
    payload = request.get_json(silent=True) or {}
    new_status = payload.get("status")
    if new_status not in _V5_VALID_STATUSES:
        return jsonify({"error": f"status must be one of {sorted(_V5_VALID_STATUSES)}"}), 400
    data["status"] = new_status
    _v5_save_skeleton(p, data)
    return jsonify({"ok": True, "status": new_status})


@app.post("/api/v5/skeleton/<skeleton_id>/overlap-decision")
def api_v5_set_overlap_decision(skeleton_id):
    p, data, err = _v5_load_skeleton(skeleton_id)
    if err:
        return jsonify({"error": err}), 404
    payload = request.get_json(silent=True) or {}
    decision = payload.get("decision")
    if decision not in _V5_VALID_OVERLAP_DECISIONS:
        return jsonify({"error": f"decision must be one of {sorted(_V5_VALID_OVERLAP_DECISIONS)}"}), 400
    data["overlap_decision"] = decision
    # "reject" decision also flips the skeleton status — it's the
    # explicit "this slide is unusable" outcome of overlap review.
    if decision == "reject":
        data["status"] = "rejected"
    _v5_save_skeleton(p, data)
    return jsonify({"ok": True, "overlap_decision": decision, "status": data["status"]})


_V5_VALID_ROLES = frozenset({
    "page_title", "subtitle", "body", "footer", "date", "page_number",
    "footnote", "caption", "key_points", "detailed_list", "cta",
    "byline", "kpi_label", "kpi_value", "section_header",
})


@app.post("/api/v5/skeleton/<skeleton_id>/slot/<slot_id>/role")
def api_v5_set_slot_role(skeleton_id, slot_id):
    """Set or clear the `role` on a slot. role=null removes the field
    (slot falls back to id/kind-based matching). Stamps user_edited
    so re-ingest doesn't blow away the override.
    """
    p, data, err = _v5_load_skeleton(skeleton_id)
    if err:
        return jsonify({"error": err}), 404
    payload = request.get_json(silent=True) or {}
    new_role = payload.get("role")
    if new_role is not None and new_role not in _V5_VALID_ROLES:
        return jsonify({"error": f"role must be null or one of {sorted(_V5_VALID_ROLES)}"}), 400
    slots = data.get("slots") or []
    target = next((s for s in slots if s.get("id") == slot_id), None)
    if target is None:
        return jsonify({"error": f"slot {slot_id!r} not found"}), 404
    if new_role is None:
        target.pop("role", None)
    else:
        target["role"] = new_role
    target["user_edited"] = True
    _v5_save_skeleton(p, data)
    return jsonify({"ok": True, "role": new_role})


@app.post("/api/v5/skeleton/<skeleton_id>/slot/<slot_id>/kind")
def api_v5_reclassify_slot_kind(skeleton_id, slot_id):
    """Change a slot's kind in-place. Resets constraints to defaults
    for the new kind and stamps user_edited:true so re-ingest preserves
    the change instead of reverting to the heuristic-derived kind.
    """
    p, data, err = _v5_load_skeleton(skeleton_id)
    if err:
        return jsonify({"error": err}), 404
    payload = request.get_json(silent=True) or {}
    new_kind = payload.get("kind")
    if new_kind not in _V5_VALID_KINDS:
        return jsonify({"error": f"kind must be one of {sorted(_V5_VALID_KINDS)}"}), 400

    slots = data.get("slots") or []
    target = next((s for s in slots if s.get("id") == slot_id), None)
    if target is None:
        return jsonify({"error": f"slot {slot_id!r} not found"}), 404
    if target.get("kind") == new_kind:
        return jsonify({"ok": True, "unchanged": True})

    # Reset constraints to the new kind's defaults — the old kind's
    # constraints don't carry meaning when the type changes.
    target["kind"] = new_kind
    target["constraints"] = _v5_default_slot_for_kind(new_kind, {}, set())["constraints"]
    target["user_edited"] = True
    _v5_save_skeleton(p, data)
    return jsonify({"ok": True, "slot": target})


@app.post("/api/v5/skeleton/<skeleton_id>/promote-shape")
def api_v5_promote_shape(skeleton_id):
    """Move an unmapped_shapes entry into slots with the chosen kind.

    Constraint defaults match what the slot builders in ingest_v5 would
    have produced for that kind — agent doesn't see this slot as
    'different' from heuristic-derived ones, just stamped with a
    user-chosen kind. shape_id is preserved so re-ingest doesn't undo
    the promotion.

    Optional `propagate: true` finds the same picture SHA in OTHER
    skeletons' unmapped_shapes and promotes there too. Use when the
    user wants consistent treatment of a deck-wide repeated picture.
    Returns counts of how many other skeletons were touched.
    """
    p, data, err = _v5_load_skeleton(skeleton_id)
    if err:
        return jsonify({"error": err}), 404
    payload = request.get_json(silent=True) or {}
    idx = payload.get("shape_index")
    kind = payload.get("kind")
    propagate = bool(payload.get("propagate"))
    if kind not in _V5_VALID_KINDS:
        return jsonify({"error": f"kind must be one of {sorted(_V5_VALID_KINDS)}"}), 400
    unmapped = data.get("unmapped_shapes") or []
    if not isinstance(idx, int) or idx < 0 or idx >= len(unmapped):
        return jsonify({"error": "shape_index out of range"}), 400
    entry = unmapped.pop(idx)
    used_ids = {s.get("id") for s in (data.get("slots") or []) if s.get("id")}
    new_slot = _v5_default_slot_for_kind(kind, entry, used_ids)
    slots = data.get("slots") or []
    slots.append(new_slot)
    data["slots"] = slots
    data["unmapped_shapes"] = unmapped
    _v5_save_skeleton(p, data)

    propagated_to: list[str] = []
    target_sha = entry.get("sha")
    if propagate and target_sha:
        for sk_dir in V5_SKELETONS_DIR.iterdir():
            if not sk_dir.is_dir() or sk_dir.name == skeleton_id:
                continue
            sk_yaml = sk_dir / "skeleton.yaml"
            if not sk_yaml.exists():
                continue
            try:
                sk_data = yaml.safe_load(sk_yaml.read_text(encoding="utf-8")) or {}
            except Exception:
                continue
            sk_unmapped = sk_data.get("unmapped_shapes") or []
            match_idx = next(
                (i for i, u in enumerate(sk_unmapped) if u.get("sha") == target_sha),
                None,
            )
            if match_idx is None:
                continue
            sib_entry = sk_unmapped.pop(match_idx)
            sk_used_ids = {s.get("id") for s in (sk_data.get("slots") or []) if s.get("id")}
            sib_slot = _v5_default_slot_for_kind(kind, sib_entry, sk_used_ids)
            sk_slots = sk_data.get("slots") or []
            sk_slots.append(sib_slot)
            sk_data["slots"] = sk_slots
            sk_data["unmapped_shapes"] = sk_unmapped
            _v5_save_skeleton(sk_yaml, sk_data)
            propagated_to.append(sk_dir.name)

    return jsonify({"ok": True, "slot": new_slot, "propagated_to": propagated_to})


@app.post("/api/v5/skeleton/<skeleton_id>/demote-slot")
def api_v5_demote_slot(skeleton_id):
    """Reverse a promote — move a user-promoted slot back to
    unmapped_shapes so it stops being addressable by the agent.

    Optional `propagate: true` finds slots in OTHER skeletons with the
    same SHA + user_promoted flag and demotes them all in one shot.
    Useful for undoing a mass promote, or cleaning up leftovers from
    earlier experimentation.
    """
    p, data, err = _v5_load_skeleton(skeleton_id)
    if err:
        return jsonify({"error": err}), 404
    payload = request.get_json(silent=True) or {}
    slot_id = payload.get("slot_id")
    propagate = bool(payload.get("propagate"))
    slots = data.get("slots") or []
    target_idx = next((i for i, s in enumerate(slots) if s.get("id") == slot_id), None)
    if target_idx is None:
        return jsonify({"error": f"slot {slot_id!r} not found"}), 404
    target = slots[target_idx]
    if not target.get("user_promoted"):
        return jsonify({
            "error": "only user_promoted slots can be demoted — heuristic-derived "
            "slots come back on re-ingest anyway",
        }), 400
    # Reconstruct an unmapped entry from the slot.
    unmapped_entry = {
        "shape_id": target.get("shape_id"),
        "kind_hint": "picture" if target.get("kind") == "image" else target.get("kind"),
        "geometry": target.get("geometry") or {},
        "skipped_reason": "user-demoted from slot",
    }
    if target.get("sha"):
        unmapped_entry["sha"] = target["sha"]
    if target.get("source_excerpt"):
        unmapped_entry["source_excerpt"] = target["source_excerpt"]
    # Apply locally.
    slots.pop(target_idx)
    unmapped = data.get("unmapped_shapes") or []
    unmapped.append(unmapped_entry)
    data["slots"] = slots
    data["unmapped_shapes"] = unmapped
    _v5_save_skeleton(p, data)

    propagated_to: list[str] = []
    target_sha = target.get("sha")
    if propagate and target_sha:
        for sk_dir in V5_SKELETONS_DIR.iterdir():
            if not sk_dir.is_dir() or sk_dir.name == skeleton_id:
                continue
            sk_yaml = sk_dir / "skeleton.yaml"
            if not sk_yaml.exists():
                continue
            try:
                sk_data = yaml.safe_load(sk_yaml.read_text(encoding="utf-8")) or {}
            except Exception:
                continue
            sk_slots = sk_data.get("slots") or []
            sib_idx = next(
                (i for i, s in enumerate(sk_slots)
                 if s.get("sha") == target_sha and s.get("user_promoted")),
                None,
            )
            if sib_idx is None:
                continue
            sib = sk_slots.pop(sib_idx)
            sib_unmapped = {
                "shape_id": sib.get("shape_id"),
                "kind_hint": "picture",
                "geometry": sib.get("geometry") or {},
                "skipped_reason": "user-demoted from slot",
                "sha": target_sha,
            }
            sk_unmapped = sk_data.get("unmapped_shapes") or []
            sk_unmapped.append(sib_unmapped)
            sk_data["slots"] = sk_slots
            sk_data["unmapped_shapes"] = sk_unmapped
            _v5_save_skeleton(sk_yaml, sk_data)
            propagated_to.append(sk_dir.name)

    return jsonify({"ok": True, "demoted": slot_id, "propagated_to": propagated_to})


def _v5_default_slot_for_kind(kind: str, unmapped_entry: dict, used_ids: set) -> dict:
    """Reasonable default constraints per kind — matches the shape the
    heuristic in ingest_v5 produces, so the agent sees no difference
    between heuristic-derived and user-promoted slots.
    """
    base_id_for_kind = {
        "heading": "heading", "paragraph": "body", "bullets": "body",
        "image": "hero", "table": "data_table", "chart": "data_chart",
        "footer": "footer",
    }
    base = base_id_for_kind.get(kind, "field")
    slot_id = base if base not in used_ids else next(
        f"{base}_{i}" for i in range(2, 99) if f"{base}_{i}" not in used_ids
    )
    constraints_for_kind = {
        "heading": {"max_chars": 60, "max_lines": 1, "required": True},
        "paragraph": {"max_chars": 200, "max_lines": 4, "required": False},
        "bullets": {"max_items": 5, "max_chars_per_item": 80, "required": False},
        "image": {"aspect": "free", "required": True, "auto_fit": "cover"},
        "table": {"max_rows": 8, "max_cols": 4, "has_header": True, "required": True},
        "chart": {"chart_type": "unknown", "max_series": 4, "max_categories": 12, "required": True},
        "footer": {"max_chars": 40, "max_lines": 1, "required": False},
    }
    excerpt = unmapped_entry.get("source_excerpt", "")
    out: dict = {
        "id": slot_id,
        "kind": kind,
        "geometry": unmapped_entry.get("geometry", {}),
        "constraints": constraints_for_kind.get(kind, {"required": False}),
        "shape_id": unmapped_entry.get("shape_id"),
        "user_promoted": True,
    }
    # Carry the SHA forward for picture promotes — demote-slot uses it
    # to find this slot's entry in unmapped_shapes after re-ingest.
    if unmapped_entry.get("sha"):
        out["sha"] = unmapped_entry["sha"]
    if excerpt:
        out["source_excerpt"] = excerpt
    return out


V5_HTML = r"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>pptx-skill v5 — skeletons</title>
  <style>
    * { box-sizing: border-box; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
           margin: 0; height: 100vh; display: flex; color: #222; }

    .sidebar { width: 280px; border-right: 1px solid #ddd; background: #fafafa;
               overflow-y: auto; display: flex; flex-direction: column; }
    .sidebar header { padding: 12px 14px; border-bottom: 1px solid #ddd; }
    .sidebar header h1 { font-size: 14px; margin: 0 0 4px; font-weight: 600; }
    .sidebar header .nav { font-size: 11px; color: #666; }
    .sidebar header .nav a { color: #0066cc; text-decoration: none; margin-right: 8px; }
    .sidebar header .nav a:hover { text-decoration: underline; }
    .deck-group h2 { font-size: 11px; text-transform: uppercase;
                     letter-spacing: 0.5px; color: #555; padding: 12px 14px 6px;
                     margin: 0; background: #f0f0f0; border-bottom: 1px solid #e0e0e0; }
    .item-list { list-style: none; padding: 0; margin: 0; }
    .item-list li { padding: 8px 14px; cursor: pointer; font-size: 12px;
                    display: flex; justify-content: space-between; align-items: center;
                    border-bottom: 1px solid #eee; gap: 8px; }
    .item-list li:hover { background: #eef4ff; }
    .item-list li.active { background: #d8e8ff; }
    .item-list .label { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
    .item-list .label .top { display: flex; align-items: center; gap: 4px; }
    .item-list .cats { font-size: 10px; color: #888; white-space: nowrap;
                       overflow: hidden; text-overflow: ellipsis; max-width: 160px; }
    .pill { font-size: 9px; padding: 2px 6px; border-radius: 8px;
            text-transform: uppercase; letter-spacing: 0.4px; font-weight: 600; flex-shrink: 0; }
    .pill.pending { background: #ffeaa7; color: #8c6900; }
    .pill.done { background: #c8e6c9; color: #1b5e20; }
    .pill.rejected { background: #e0e0e0; color: #555; }
    .warn-dot { color: #d68b00; font-weight: 700; }

    .preview-area { flex: 1; background: #1c1c1c; display: flex; align-items: center;
                    justify-content: center; padding: 20px; min-width: 0; }
    .preview-wrap { position: relative; max-width: 100%; max-height: 100%;
                    background: white; box-shadow: 0 4px 30px rgba(0,0,0,0.4); }
    .preview-wrap img { display: block; max-width: 100%;
                        max-height: calc(100vh - 80px); object-fit: contain; }
    .slot-overlay { position: absolute; border: 2px solid;
                    cursor: pointer; transition: background 0.1s; }
    .slot-overlay:hover { background: rgba(255,255,255,0.18) !important; }
    .slot-overlay .slot-label { position: absolute; top: 2px; left: 2px;
                                background: rgba(0,0,0,0.85); color: white;
                                padding: 2px 6px; font-size: 10px; border-radius: 3px;
                                font-family: ui-monospace, monospace; white-space: nowrap; }
    .slot-overlay.heading   { border-color: #e53935; }
    .slot-overlay.paragraph { border-color: #fb8c00; }
    .slot-overlay.bullets   { border-color: #1e88e5; }
    .slot-overlay.table     { border-color: #43a047; }
    .slot-overlay.chart     { border-color: #8e24aa; }
    .slot-overlay.image     { border-color: #00acc1; }
    .slot-overlay.footer    { border-color: #757575; }
    .slot-overlay.unmapped  { border-color: #999; border-style: dashed; opacity: 0.7; }
    .slot-overlay.unmapped .slot-label { background: rgba(80,80,80,0.85); }

    .preview-empty { color: #999; font-size: 13px; text-align: center; padding: 40px;
                     line-height: 1.5; }
    .preview-empty .hint { font-size: 11px; color: #777; margin-top: 8px;
                           font-family: ui-monospace, monospace; }

    .panel { width: 440px; border-left: 1px solid #ddd; padding: 16px;
             overflow-y: auto; background: white; }
    .panel h2 { margin: 0 0 4px; font-size: 15px; word-break: break-all;
                font-family: ui-monospace, monospace; }
    .panel .subtitle { color: #777; font-size: 12px; margin-bottom: 14px; }
    .panel .cats-row { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 14px;
                       align-items: center; }
    .cats-row .label { font-size: 11px; color: #555; text-transform: uppercase;
                        letter-spacing: 0.5px; margin-right: 4px; }
    .cat-tag { background: #e8eaf6; color: #3949ab; padding: 3px 8px;
               border-radius: 10px; font-size: 11px; font-weight: 500; }
    .warning-banner { background: #fff4e0; border: 1px solid #ffcc80; color: #8c5400;
                      padding: 8px 12px; border-radius: 4px; font-size: 12px;
                      margin-bottom: 14px; line-height: 1.4; }

    .slots-header { font-size: 11px; text-transform: uppercase;
                    letter-spacing: 0.5px; color: #555; margin: 0 0 8px; }
    .slot-card { border: 1px solid #e0e0e0; border-radius: 4px; padding: 10px 12px;
                 margin-bottom: 10px; background: #fafafa; transition: background 0.1s; }
    .slot-card.active { background: #fff8e1; border-color: #ffc107; }
    .slot-head { display: flex; justify-content: space-between;
                 align-items: center; margin-bottom: 6px; gap: 8px; }
    .slot-id { font-weight: 600; font-size: 13px; font-family: ui-monospace, monospace; }
    .kind-btns { display: flex; gap: 2px; flex-wrap: wrap; }
    .kind-btn { padding: 2px 6px; border: 1px solid #ccc; background: white;
                font-size: 10px; cursor: pointer; border-radius: 3px;
                color: #555; font-family: ui-monospace, monospace; }
    .kind-btn:hover { background: #eef4ff; color: #0066cc; border-color: #0066cc; }
    .kind-btn.active { background: #1e88e5; color: white; border-color: #1565c0;
                       cursor: default; font-weight: 600; }
    .kind-btn.active:hover { background: #1e88e5; color: white; border-color: #1565c0; }
    .user-edited-flag { color: #1565c0; font-size: 10px; font-style: italic;
                         margin-left: 6px; }
    .role-badge { display: inline-block; font-size: 10px; padding: 1px 6px;
                  background: #ede7f6; color: #5e35b1; border-radius: 8px;
                  font-family: ui-monospace, monospace; margin-left: 6px;
                  cursor: pointer; }
    .role-badge:hover { background: #d1c4e9; }
    .role-badge.unset { background: #f5f5f5; color: #999; }
    .excerpt { font-size: 12px; color: #444; margin-bottom: 6px;
               font-style: italic; word-break: break-word; line-height: 1.4; }
    .constraints { font-size: 11px; color: #666;
                   font-family: ui-monospace, monospace; }
    .constraints .req { color: #d32f2f; font-weight: 600; }
    .constraints .sep { color: #bbb; margin: 0 4px; }

    .actions-row { display: flex; gap: 6px; margin: 8px 0 14px; flex-wrap: wrap; }
    .action-btn { padding: 4px 10px; border: 1px solid #ccc; background: white;
                  font-size: 11px; cursor: pointer; border-radius: 3px;
                  font-family: inherit; }
    .action-btn:hover { background: #f0f4ff; border-color: #0066cc; color: #0066cc; }
    .action-btn.danger:hover { background: #ffeaea; border-color: #d32f2f; color: #d32f2f; }
    .action-btn.primary { background: #1e88e5; color: white; border-color: #1565c0; }
    .action-btn.primary:hover { background: #1565c0; color: white; }
    .action-btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .action-btn:disabled:hover { background: white; border-color: #ccc; color: #888; }

    .overlap-banner { background: #fff4e0; border: 1px solid #ffcc80; color: #8c5400;
                      padding: 10px 12px; border-radius: 4px; font-size: 12px;
                      margin-bottom: 14px; line-height: 1.4; }
    .overlap-banner .head { font-weight: 600; margin-bottom: 6px; }
    .overlap-banner .actions { display: flex; gap: 6px; margin-top: 8px; flex-wrap: wrap; }

    .unmapped-section { margin-top: 18px; padding-top: 14px;
                        border-top: 1px solid #e0e0e0; }
    .unmapped-section h3 { font-size: 11px; text-transform: uppercase;
                           letter-spacing: 0.5px; color: #555; margin: 0 0 8px;
                           display: flex; justify-content: space-between;
                           align-items: center; cursor: pointer; user-select: none; }
    .unmapped-section h3 .toggle { color: #999; font-weight: normal; font-size: 12px; }
    .unmapped-card { border: 1px dashed #bbb; border-radius: 4px; padding: 8px 12px;
                     margin-bottom: 8px; background: #f9f9f9; }
    .unmapped-card.active { background: #f0f4ff; border-color: #0066cc; }
    .unmapped-card .hint-head { display: flex; justify-content: space-between;
                                align-items: center; margin-bottom: 4px; gap: 6px; }
    .unmapped-card .kind-hint { font-family: ui-monospace, monospace; font-size: 11px;
                                color: #555; background: #e8e8e8; padding: 1px 5px;
                                border-radius: 2px; }
    .unmapped-card .reason { font-size: 10px; color: #888; font-style: italic; }
    .unmapped-card .excerpt { font-size: 12px; color: #444; margin: 4px 0;
                              word-break: break-word; }
    .unmapped-card .promote-row { display: flex; gap: 3px; flex-wrap: wrap; margin-top: 4px;
                                  align-items: center; }
    .unmapped-card .promote-label { font-size: 10px; color: #666; margin-right: 4px; }
    .promote-btn { padding: 2px 7px; border: 1px solid #aaa; background: white;
                   font-size: 10px; cursor: pointer; border-radius: 3px;
                   font-family: ui-monospace, monospace; color: #333; }
    .promote-btn:hover { background: #1e88e5; color: white; border-color: #1565c0; }

    /* Belt-and-suspenders: hide v4 debug-floater if it leaks in. */
    .dbg-floater { display: none !important; }
  </style>
</head>
<body>
  <aside class="sidebar">
    <header>
      <h1>pptx-skill v5 — skeletons</h1>
      <div class="nav"><a href="/">describe</a><a href="/compose">compose</a></div>
    </header>
    <div id="skeletons-list"></div>
  </aside>
  <main class="preview-area" id="preview-area">
    <div class="preview-empty">
      <div>Select a skeleton on the left.</div>
      <div class="hint">colored boxes = our digest proposal · hover for details</div>
    </div>
  </main>
  <aside class="panel" id="panel">
    <div class="preview-empty" style="text-align:center;padding-top:40px;">
      No skeleton selected.
    </div>
  </aside>

  <script>
    const KIND_LIST = ['heading', 'paragraph', 'bullets', 'table', 'chart', 'image', 'footer'];
    const ROLE_LIST = ['(none)', 'page_title', 'subtitle', 'body', 'footer', 'date',
      'page_number', 'footnote', 'caption', 'key_points', 'detailed_list',
      'cta', 'byline', 'kpi_label', 'kpi_value', 'section_header'];

    async function loadSkeletons() {
      const r = await fetch('/api/v5/skeletons');
      const data = await r.json();
      const root = document.getElementById('skeletons-list');
      root.innerHTML = '';
      const decks = Object.keys(data.decks).sort();
      if (decks.length === 0) {
        root.innerHTML = '<div style="padding:14px;color:#888;font-size:12px;line-height:1.5;">No skeletons yet.<br>Run:<br><code style="font-size:11px;">python3 authoring/cli.py ingest your_deck.pptx</code></div>';
        return;
      }
      decks.forEach(deck => {
        const group = document.createElement('div');
        group.className = 'deck-group';
        const h = document.createElement('h2');
        h.textContent = deck;
        group.appendChild(h);
        const ul = document.createElement('ul');
        ul.className = 'item-list';
        data.decks[deck].forEach(sk => {
          const li = document.createElement('li');
          li.dataset.id = sk.id;
          const cats = (sk.categories || []).join(', ') || '—';
          const warn = sk.has_warnings ? '<span class="warn-dot" title="overlap_detected">⚠</span> ' : '';
          li.innerHTML = `
            <span class="label">
              <span class="top">${warn}slide ${sk.source_slide_index} · ${sk.slot_count} slots</span>
              <span class="cats">${cats}</span>
            </span>
            <span class="pill ${sk.status}">${sk.status}</span>
          `;
          li.onclick = () => selectSkeleton(sk.id);
          ul.appendChild(li);
        });
        group.appendChild(ul);
        root.appendChild(group);
      });
    }

    async function selectSkeleton(id) {
      document.querySelectorAll('.item-list li').forEach(li => {
        li.classList.toggle('active', li.dataset.id === id);
      });
      const r = await fetch(`/api/v5/skeleton/${encodeURIComponent(id)}`);
      const sk = await r.json();
      renderPreview(sk);
      renderPanel(sk);
    }

    function renderPreview(sk) {
      const area = document.getElementById('preview-area');
      area.innerHTML = '';
      const wrap = document.createElement('div');
      wrap.className = 'preview-wrap';
      const img = document.createElement('img');
      img.src = `/v5/skeleton/${encodeURIComponent(sk.id)}/preview.png?cb=${Date.now()}`;
      img.onerror = () => {
        wrap.style.background = 'transparent';
        wrap.style.boxShadow = 'none';
        wrap.innerHTML = `<div class="preview-empty">
          <div>No preview rendered for this skeleton yet.</div>
          <div class="hint">Run: python3 authoring/cli.py preview</div>
          <div class="hint" style="margin-top:14px;">Slot inventory is still visible on the right →</div>
        </div>`;
      };
      img.onload = () => {
        (sk.slots || []).forEach(slot => {
          const g = slot.geometry || {};
          const d = document.createElement('div');
          d.className = `slot-overlay ${slot.kind}`;
          d.style.left = `${(g.x || 0) * 100}%`;
          d.style.top = `${(g.y || 0) * 100}%`;
          d.style.width = `${(g.w || 0) * 100}%`;
          d.style.height = `${(g.h || 0) * 100}%`;
          d.dataset.slotId = slot.id;
          d.innerHTML = `<span class="slot-label">${slot.id} · ${slot.kind}</span>`;
          d.onmouseenter = () => highlightSlot(slot.id, true);
          d.onmouseleave = () => highlightSlot(slot.id, false);
          wrap.appendChild(d);
        });
        (sk.unmapped_shapes || []).forEach((entry, idx) => {
          const g = entry.geometry || {};
          const d = document.createElement('div');
          d.className = 'slot-overlay unmapped';
          d.style.left = `${(g.x || 0) * 100}%`;
          d.style.top = `${(g.y || 0) * 100}%`;
          d.style.width = `${(g.w || 0) * 100}%`;
          d.style.height = `${(g.h || 0) * 100}%`;
          d.dataset.unmappedIndex = idx;
          d.innerHTML = `<span class="slot-label">unmapped · ${entry.kind_hint}</span>`;
          d.onmouseenter = () => highlightUnmapped(idx, true);
          d.onmouseleave = () => highlightUnmapped(idx, false);
          wrap.appendChild(d);
        });
      };
      wrap.appendChild(img);
      area.appendChild(wrap);
    }

    function renderPanel(sk) {
      const panel = document.getElementById('panel');
      const cats = (sk.categories || []).map(c => `<span class="cat-tag">${escapeHtml(c)}</span>`).join('');

      // Status action row depends on current status.
      const status = sk.status || 'pending';
      let statusActions = '';
      if (status === 'pending') {
        statusActions = `
          <button class="action-btn primary" onclick="setStatus('${sk.id}', 'done')">Mark done</button>
          <button class="action-btn danger" onclick="setStatus('${sk.id}', 'rejected')">Reject</button>`;
      } else if (status === 'done') {
        statusActions = `
          <button class="action-btn" onclick="setStatus('${sk.id}', 'pending')">Back to pending</button>
          <button class="action-btn danger" onclick="setStatus('${sk.id}', 'rejected')">Reject</button>`;
      } else if (status === 'rejected') {
        statusActions = `
          <button class="action-btn primary" onclick="setStatus('${sk.id}', 'pending')">Restore</button>`;
      }

      // Overlap banner — actionable only if no decision recorded yet.
      const warnings = sk.digest_warnings || [];
      const overlapDecided = !!sk.overlap_decision;
      let warnHtml = '';
      if (warnings.length && !overlapDecided) {
        warnHtml = `
          <div class="overlap-banner">
            <div class="head">⚠ overlap_detected</div>
            <div>A picture sits under text on this slide. The agent shouldn't blindly swap it (cross-slide misalignment risk). Choose:</div>
            <div class="actions">
              <button class="action-btn" onclick="setOverlapDecision('${sk.id}', 'image_slot')">Keep as image slot</button>
              <button class="action-btn" disabled title="B4-render not yet implemented">Freeze as background</button>
              <button class="action-btn danger" onclick="setOverlapDecision('${sk.id}', 'reject')">Reject slide</button>
            </div>
          </div>`;
      } else if (overlapDecided) {
        warnHtml = `<div class="overlap-banner" style="background:#e8f5e9;border-color:#a5d6a7;color:#1b5e20;">
          <div class="head">✓ overlap decision: ${escapeHtml(sk.overlap_decision)}</div>
        </div>`;
      }

      const slotCards = (sk.slots || []).map(slot => {
        const kindBtns = KIND_LIST.map(k =>
          `<button class="kind-btn ${k === slot.kind ? 'active' : ''}"
                   onclick="reclassifySlot('${sk.id}', '${escapeHtml(slot.id)}', '${k}')"
                   title="${k === slot.kind ? 'current kind' : 'click to reclassify to ' + k}">
             ${k}
           </button>`
        ).join('');
        const c = slot.constraints || {};
        const parts = [];
        if (c.max_chars) parts.push(`max ${c.max_chars} chars`);
        if (c.max_lines) parts.push(`${c.max_lines} lines`);
        if (c.max_items) parts.push(`${c.max_items} items`);
        if (c.max_chars_per_item) parts.push(`${c.max_chars_per_item} chars/item`);
        if (c.max_rows && c.max_cols) parts.push(`${c.max_rows}×${c.max_cols}${c.has_header ? ' +hdr' : ''}`);
        if (c.chart_type) parts.push(`${c.chart_type}, ${c.max_series}s×${c.max_categories}c`);
        if (c.aspect) parts.push(`aspect ${c.aspect}`);
        if (c.required) parts.push('<span class="req">required</span>');
        const editedFlag = slot.user_edited
          ? '<span class="user-edited-flag">(user-edited)</span>'
          : (slot.user_promoted ? '<span class="user-edited-flag">(user-promoted)</span>' : '');
        const demoteBtn = slot.user_promoted
          ? `<button class="action-btn danger" style="margin-left:6px;font-size:10px;padding:2px 6px;"
                     onclick="demoteSlot('${sk.id}', '${escapeHtml(slot.id)}', ${JSON.stringify(slot.sha || null)})"
                     title="Move this slot back to unmapped decoration">demote</button>`
          : '';
        const roleLabel = slot.role || 'no role';
        const roleClass = slot.role ? '' : 'unset';
        const roleBadge = `<span class="role-badge ${roleClass}" title="Click to change role"
                                  onclick="changeRole('${sk.id}', '${escapeHtml(slot.id)}', '${escapeHtml(slot.role || '')}')">
                              ${escapeHtml(roleLabel)}
                           </span>`;
        return `<div class="slot-card" data-slot-id="${escapeHtml(slot.id)}">
          <div class="slot-head">
            <span class="slot-id">${escapeHtml(slot.id)}${roleBadge}${editedFlag}${demoteBtn}</span>
            <span class="kind-btns">${kindBtns}</span>
          </div>
          ${slot.source_excerpt ? `<div class="excerpt">${escapeHtml(slot.source_excerpt)}</div>` : ''}
          <div class="constraints">${parts.join('<span class="sep">·</span>')}</div>
        </div>`;
      }).join('');

      const unmapped = sk.unmapped_shapes || [];
      const unmappedHtml = unmapped.length ? `
        <div class="unmapped-section">
          <h3>Unmapped shapes <span class="toggle">(${unmapped.length}) — promote any our heuristic missed</span></h3>
          ${unmapped.map((entry, idx) => {
            const promoteBtns = KIND_LIST.map(k =>
              `<button class="promote-btn" onclick="promoteShape('${sk.id}', ${idx}, '${k}')">${k}</button>`
            ).join('');
            return `<div class="unmapped-card" data-unmapped-index="${idx}">
              <div class="hint-head">
                <span class="kind-hint">${escapeHtml(entry.kind_hint)}</span>
                <span class="reason">${escapeHtml(entry.skipped_reason || '')}</span>
              </div>
              ${entry.source_excerpt
                ? `<div class="excerpt">${escapeHtml(entry.source_excerpt)}</div>`
                : '<div class="excerpt" style="color:#999;">(no text)</div>'}
              <div class="promote-row">
                <span class="promote-label">promote to →</span>
                ${promoteBtns}
              </div>
            </div>`;
          }).join('')}
        </div>` : '';

      panel.innerHTML = `
        <h2>${escapeHtml(sk.id)}</h2>
        <div class="subtitle">
          ${escapeHtml(sk.source_deck)} · slide ${sk.source_slide_index}
          · ${(sk.slots || []).length} slot${(sk.slots || []).length === 1 ? '' : 's'}
          · <span class="pill ${status}">${status}</span>
        </div>
        <div class="actions-row">${statusActions}</div>
        <div class="cats-row"><span class="label">categories</span>${cats || '<span style="color:#888;font-size:11px;">none</span>'}</div>
        ${warnHtml}
        <h3 class="slots-header">slots (proposed)</h3>
        ${slotCards}
        ${unmappedHtml}
      `;
    }

    async function setStatus(id, status) {
      const r = await fetch(`/api/v5/skeleton/${encodeURIComponent(id)}/status`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({status})
      });
      if (!r.ok) {
        alert('Status update failed: ' + (await r.text()));
        return;
      }
      await Promise.all([loadSkeletons(), refreshCurrent(id)]);
    }

    async function setOverlapDecision(id, decision) {
      const r = await fetch(`/api/v5/skeleton/${encodeURIComponent(id)}/overlap-decision`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({decision})
      });
      if (!r.ok) { alert('Overlap decision failed: ' + (await r.text())); return; }
      await Promise.all([loadSkeletons(), refreshCurrent(id)]);
    }

    async function promoteShape(id, shape_index, kind) {
      // If the target unmapped entry is a deck-wide repeated brand mark,
      // ask whether to propagate so the user doesn't accidentally create
      // a slide-N-only swappable hero while the same picture stays
      // decoration on every other slide. The /v5 panel already has the
      // skeleton in scope; we look up the entry from the cached render.
      const panel = document.getElementById('panel');
      const card = panel.querySelector(`.unmapped-card[data-unmapped-index="${shape_index}"]`);
      let propagate = false;
      if (card) {
        const reason = card.querySelector('.reason')?.textContent || '';
        if (reason.toLowerCase().includes('repeated brand mark')) {
          const choice = confirm(
            "This picture appears on multiple slides as deck-wide decoration.\n\n" +
            "OK   = promote on all slides where it appears (consistent)\n" +
            "Cancel = promote on this slide only (may cause cross-slide inconsistency)"
          );
          propagate = choice === true;
        }
      }
      const r = await fetch(`/api/v5/skeleton/${encodeURIComponent(id)}/promote-shape`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({shape_index, kind, propagate})
      });
      if (!r.ok) { alert('Promote failed: ' + (await r.text())); return; }
      const result = await r.json();
      if (result.propagated_to?.length) {
        // Refresh the sidebar list so other skeletons show updated state.
        await loadSkeletons();
      }
      await refreshCurrent(id);
    }

    async function demoteSlot(id, slot_id, sha) {
      let propagate = false;
      if (sha) {
        propagate = confirm(
          "Demote this slot back to decoration.\n\n" +
          "OK   = demote on all slides where the same picture is user-promoted\n" +
          "Cancel = demote on this slide only"
        );
      }
      const r = await fetch(`/api/v5/skeleton/${encodeURIComponent(id)}/demote-slot`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({slot_id, propagate})
      });
      if (!r.ok) { alert('Demote failed: ' + (await r.text())); return; }
      const result = await r.json();
      if (result.propagated_to?.length) {
        await loadSkeletons();
      }
      await refreshCurrent(id);
    }

    async function reclassifySlot(id, slot_id, kind) {
      const r = await fetch(`/api/v5/skeleton/${encodeURIComponent(id)}/slot/${encodeURIComponent(slot_id)}/kind`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({kind})
      });
      if (!r.ok) { alert('Reclassify failed: ' + (await r.text())); return; }
      await refreshCurrent(id);
    }

    async function changeRole(id, slot_id, currentRole) {
      const prompt_text = "Set slot role (blank = clear). Allowed:\n" +
        ROLE_LIST.slice(1).join(', ');
      const answer = window.prompt(prompt_text, currentRole || '');
      if (answer === null) return;  // cancelled
      const new_role = answer.trim() === '' ? null : answer.trim();
      const r = await fetch(`/api/v5/skeleton/${encodeURIComponent(id)}/slot/${encodeURIComponent(slot_id)}/role`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({role: new_role})
      });
      if (!r.ok) { alert('Role update failed: ' + (await r.text())); return; }
      await refreshCurrent(id);
    }

    async function refreshCurrent(id) {
      const r = await fetch(`/api/v5/skeleton/${encodeURIComponent(id)}`);
      const sk = await r.json();
      renderPreview(sk);
      renderPanel(sk);
    }

    function highlightUnmapped(idx, on) {
      document.querySelectorAll(`.unmapped-card[data-unmapped-index="${idx}"]`).forEach(c => {
        c.classList.toggle('active', on);
      });
    }

    function highlightSlot(slotId, on) {
      document.querySelectorAll(`.slot-card[data-slot-id="${CSS.escape(slotId)}"]`).forEach(c => {
        c.classList.toggle('active', on);
      });
    }

    function escapeHtml(s) {
      return (s == null ? '' : String(s)).replace(/[&<>"']/g,
        c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    }

    loadSkeletons();
  </script>
</body>
</html>
"""


def main():
    port = int(os.environ.get("PPTX_SKILL_PORT", "5050"))
    # On the v5 redesign branch, /v5 is the primary surface — auto-open
    # there rather than the v4 describe page. Pass PPTX_SKILL_LANDING=/
    # to keep the v4 landing for asset-describe sessions.
    landing = os.environ.get("PPTX_SKILL_LANDING", "/v5")
    url = f"http://127.0.0.1:{port}{landing}"
    Timer(1.2, lambda: webbrowser.open(url)).start()
    print(f"pptx-skill app → {url}")
    print(f"  v5 skeletons:   http://127.0.0.1:{port}/v5")
    print(f"  v4 describe:    http://127.0.0.1:{port}/")
    print(f"  compose:        http://127.0.0.1:{port}/compose")
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
