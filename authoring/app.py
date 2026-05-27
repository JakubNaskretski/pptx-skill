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
from flask import Flask, abort, g, jsonify, redirect, render_template_string, request, send_file

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
    # All sidecars under workspace/ are now asset records (v4 slide
    # YAMLs are no longer generated). Function kept as the explicit
    # kind hook for the describe UI / save path.
    return "asset"


def _items() -> dict:
    out: dict = {"assets": []}
    for p in cli_mod.iter_asset_yamls():
        d = _safe_read(p)
        out["assets"].append({
            "id": d.get("id") or p.stem,
            "yaml": str(p.relative_to(HERE)),
            "status": d.get("status", "pending"),
        })
    return out


def _asset_binary(yaml_path: Path) -> Path | None:
    for cand in yaml_path.parent.glob(f"{yaml_path.stem}.*"):
        if cand.suffix != ".yaml":
            return cand
    return None


ASSET_DESCRIPTIVE = ("kind", "tags", "description", "notes")
_LIST_KEYS = {"tags"}


def _descriptive_yaml(data: dict, kind: str) -> str:
    subset: dict = {}
    for k in ASSET_DESCRIPTIVE:
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
    errs = cli_mod.validate_asset(existing)
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


@app.post("/api/asset/add")
def api_add_asset():
    """Register a single image / SVG / XML fragment as a library asset.

    Form fields:
      file (required) — the binary
      kind (optional) — override the inferred kind (must match the
                        asset-vocab enum)

    Returns ``{"asset_id", "sha1", "yaml_path", "binary_path", "kind"}``
    on success. Idempotent — uploading the same content twice returns
    the existing asset's id without rewriting the YAML.
    """
    if "file" not in request.files:
        return jsonify({"error": "no file uploaded (expect form field 'file')"}), 400
    f = request.files["file"]
    name = f.filename or ""
    if not name:
        return jsonify({"error": "uploaded file has no name"}), 400
    ext = Path(name).suffix.lower()
    if ext not in cli_mod._ADD_ASSET_EXT_KIND:
        return jsonify({
            "error": (
                f"unsupported extension {ext!r}; expected one of "
                f"{sorted(cli_mod._ADD_ASSET_EXT_KIND)}"
            ),
        }), 400
    kind_hint = (request.form.get("kind") or "").strip() or None
    if kind_hint and kind_hint not in cli_mod.ASSET_KIND_ENUM:
        return jsonify({
            "error": (
                f"invalid kind {kind_hint!r}; expected one of "
                f"{sorted(cli_mod.ASSET_KIND_ENUM)}"
            ),
        }), 400
    safe_basename = Path(name).name.lstrip(".").strip() or f"upload{ext}"
    with tempfile.TemporaryDirectory(prefix="pptx_asset_") as td:
        tmp_path = Path(td) / safe_basename
        f.save(str(tmp_path))
        try:
            entry = cli_mod._add_asset_to_workspace(tmp_path, kind_hint=kind_hint)
        except Exception as e:
            debug_event("error", "add-asset",
                        f"add-asset failed for {safe_basename}: {type(e).__name__}: {e}")
            return jsonify({"error": f"add-asset failed: {e}"}), 500
    debug_event(
        "info", "add-asset",
        f"added {entry['asset_id']} (kind={entry['kind'] or 'unset'}) from {safe_basename}",
        **entry,
    )
    return jsonify(entry)


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
    text = (HERE / "prompts" / "describe_asset.md").read_text(encoding="utf-8")
    return jsonify({"text": text})


# ---------------------------------------------------------------------------
# Batch describe
# ---------------------------------------------------------------------------


def _bulk_instructions(kind: str, n: int, per_item_prompt: str) -> str:
    """Wrapper prompt for a bulk-describe batch. Tells the model the
    expected JSON shape and inlines the per-item schema. kind is always
    "asset" today — kept as a parameter so the caller stays explicit.
    """
    del kind  # currently asset-only; reserved for future kinds
    sample_block = (
        '{\n'
        '  "01": {\n'
        '    "kind": "photo",\n'
        '    "tags": ["people", "office"],\n'
        '    "description": "...",\n'
        '    "notes": ""\n'
        '  },\n'
        '  "02": { "kind": "photo", "...": "same fields" }\n'
        '}\n'
    )
    return (
        f"# Bulk describe batch — {n} image(s)\n\n"
        f"You will see {n} images numbered 01 through {n:02d}. For each, "
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
    try:
        count = max(1, min(20, int(body.get("count", 10))))
    except (TypeError, ValueError):
        return jsonify({"error": "count must be an integer"}), 400

    candidates = list(cli_mod.iter_asset_yamls())
    pending = [p for p in candidates if _safe_read(p).get("status", "pending") == "pending"]
    selected = pending[:count]
    if not selected:
        return jsonify({"error": "no pending assets"}), 400

    BATCHES_DIR.mkdir(exist_ok=True)
    batch_id = datetime.now().strftime("%Y%m%d-%H%M%S")
    batch_dir = BATCHES_DIR / batch_id
    batch_dir.mkdir(exist_ok=True)

    manifest: dict = {"kind": "asset", "created": batch_id, "items": {}}
    skipped: list[dict] = []
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        added = 0
        for ypath in selected:
            rel = str(ypath.relative_to(HERE))
            try:
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
            except OSError as e:
                skipped.append({"yaml": rel, "reason": f"OS error: {e}"})
            except Exception as e:
                skipped.append({"yaml": rel, "reason": f"unexpected: {type(e).__name__}: {e}"})

        per_item_prompt = (HERE / "prompts" / "describe_asset.md").read_text(encoding="utf-8")
        zf.writestr(
            "instructions.md",
            _bulk_instructions("asset", added, per_item_prompt),
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
        f"describe batch {batch_id} created — "
        f"{added}/{count} assets, {len(skipped)} skipped"
        + (f", pruned {pruned} old" if pruned else ""),
        batch_id=batch_id, count=added, requested=count,
        skipped=len(skipped),
    )

    return jsonify({
        "batch_id": batch_id,
        "kind": "asset",
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

    results = []
    for key, manifest_entry in manifest["items"].items():
        n = _normalize_key_to_int(key)
        entry = by_int.get(n) if n is not None else None

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
    binary = _asset_binary(p)
    if binary is None:
        abort(404, "asset binary missing")
    return send_file(binary)


# ---------------------------------------------------------------------------
# Compose flow — filter KB, build prompt bundle, run compose from a plan
# ---------------------------------------------------------------------------


FILTER_DIMENSIONS = {
    "skeletons": ("categories",),
    "assets": ("kind", "tags"),
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


def _collect_skeletons() -> list[dict]:
    """Read every workspace skeleton sidecar that's not rejected.

    Each returned dict carries `_dir` (the on-disk directory name) so
    bundle builders can re-resolve sibling files (preview.png,
    background.png) without re-globbing the workspace.
    """
    root = cli_mod.WORKSPACE / "skeletons"
    if not root.exists():
        return []
    out: list[dict] = []
    for d in sorted(root.iterdir()):
        if not d.is_dir():
            continue
        sk_path = d / "skeleton.yaml"
        if not sk_path.exists():
            continue
        data = _safe_read(sk_path)
        if not data or data.get("status") == "rejected":
            continue
        data["_dir"] = d.name
        data["_yaml_path"] = sk_path
        out.append(data)
    return out


def _collect_descriptions() -> tuple[list[dict], list[dict]]:
    """Return (skeletons, assets) for filter UI + brief bundle.

    Slides (v4) are not collected — the compose flow is v5-native; v4
    `decks/<deck>/slides/*.yaml` records are ignored here even if they
    still exist on disk from a pre-redesign workspace.
    """
    skeletons = _collect_skeletons()
    assets: list[dict] = []
    for p in cli_mod.iter_asset_yamls():
        d = _safe_read(p)
        if not d:
            continue
        d["_yaml_path"] = p
        assets.append(d)
    return skeletons, assets


def _collect_filter_options() -> dict:
    skeletons, assets = _collect_descriptions()

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
        "skeletons": {
            "options": collect(skeletons, FILTER_DIMENSIONS["skeletons"]),
            "total": len(skeletons),
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
    skeletons, assets = _collect_descriptions()
    sk_filters = filters.get("skeletons") or {}
    ast_filters = filters.get("assets") or {}
    skeletons_out = [s for s in skeletons if _matches_filters(s, sk_filters)]
    assets_out = [a for a in assets if _matches_filters(a, ast_filters)]
    return skeletons_out, assets_out


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
    """Slim brief.md for the v5 bundle. Just the user's brief plus the
    user-assets section when present. The v4 four-pass preamble +
    helpers docs + slot-polymorphism notes are dropped — SKILL_v5.md
    is the canonical contract for v5 and covers the workflow.
    """
    user_section = _format_user_assets_section(user_meta or {})
    out = (
        "# Deck brief\n\n"
        f"{brief.strip() or '(no brief supplied)'}\n"
    )
    if user_section:
        out += "\n---\n\n" + user_section
    return out


def _skeleton_index_entry(sk: dict) -> dict:
    slots = sk.get("slots") or []
    return {
        "id": sk["id"],
        "source_deck": sk.get("source_deck"),
        "source_slide_index": sk.get("source_slide_index"),
        "categories": sk.get("categories") or [],
        "slot_count": len(slots),
        "slot_kinds": sorted({s.get("kind") for s in slots if s.get("kind")}),
        "status": sk.get("status", "pending"),
    }


def _v5_asset_summary(a: dict) -> dict:
    out = {
        "id": a["id"],
        "kind": a.get("kind", ""),
        "tags": a.get("tags", []),
        "description": a.get("description", ""),
    }
    for k in ("width", "height", "aspect", "colors_hex",
              "recolor_targets", "table", "chart", "shape", "smartart",
              "interpretation"):
        v = a.get(k)
        if v:
            out[k] = v
    return out


def _v5_theme_entry(theme: dict) -> dict:
    return {
        "id": theme.get("id"),
        "palette": theme.get("palette", {}),
        "fonts": theme.get("fonts", {}),
    }


def _build_v5_index(skeletons: list[dict], assets: list[dict],
                    themes: list[dict]) -> dict:
    return {
        "version": 5,
        "themes": [_v5_theme_entry(t) for t in themes],
        "skeletons": [_skeleton_index_entry(s) for s in skeletons],
        "assets": [_v5_asset_summary(a) for a in assets],
        "tag_vocab": cli_mod.load_tag_vocab(),
    }


def _build_prompt_bundle_zip(skeletons: list[dict], assets: list[dict], brief: str) -> bytes:
    """v5 brief bundle. Ships skill-v5 layout (SKILL_v5.md as SKILL.md,
    skeletons/, themes/, assets/, reader.py, tag_vocab.yaml, index.json)
    plus brand.md, brief.md, optional user_assets/ low-res previews +
    manifest. Filter narrows which skeletons and assets are included;
    themes are auto-restricted to the decks the included skeletons came
    from.
    """
    user_meta = _read_user_meta(USER_STAGED_DIR)
    skeletons_root = cli_mod.WORKSPACE / "skeletons"
    themes_root = cli_mod.WORKSPACE / "themes"

    # Only ship themes whose deck has at least one skeleton in the
    # bundle — mirrors the v4 theme-trim logic so the agent's view
    # matches what it can actually use.
    decks_in_bundle = {sk.get("source_deck") for sk in skeletons}
    decks_in_bundle.discard(None)
    themes: list[dict] = []
    if themes_root.exists():
        for d in sorted(themes_root.iterdir()):
            if not d.is_dir() or d.name not in decks_in_bundle:
                continue
            theme_yaml = d / "theme.yaml"
            if not theme_yaml.exists():
                continue
            t = _safe_read(theme_yaml)
            if not t:
                continue
            t["_dir"] = d.name
            themes.append(t)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # Consumer contract + reader so the agent can call find-asset / etc.
        zf.writestr("SKILL.md", (cli_mod.CONSUMER / "SKILL_v5.md").read_text(encoding="utf-8"))
        zf.writestr("reader.py", (cli_mod.CONSUMER / "reader.py").read_text(encoding="utf-8"))
        brand = _read_brand().strip()
        if brand:
            zf.writestr("brand.md", brand + "\n")
        zf.writestr("brief.md", _format_brief(brief, user_meta))
        zf.writestr(
            "tag_vocab.yaml",
            yaml.safe_dump({"tags": cli_mod.load_tag_vocab()},
                           sort_keys=False, allow_unicode=True),
        )

        # Themes
        for t in themes:
            tid = t.get("id") or t["_dir"]
            src_dir = themes_root / t["_dir"]
            for fn in ("theme.yaml", "master.pptx", "preview.png"):
                src = src_dir / fn
                if src.exists():
                    zf.write(src, f"themes/{tid}/{fn}")

        # Skeletons (sidecar + previews)
        for sk in skeletons:
            sid = sk["id"]
            src_dir = skeletons_root / sk["_dir"]
            for fn in ("skeleton.yaml", "preview.png", "background.png"):
                src = src_dir / fn
                if src.exists():
                    zf.write(src, f"skeletons/{sid}/{fn}")

        # KB assets (slim sidecar + binary)
        for a in assets:
            aid = a["id"]
            yaml_path: Path = a["_yaml_path"]
            clean = {k: v for k, v in a.items() if not k.startswith("_")}
            zf.writestr(
                f"assets/{aid}.yaml",
                yaml.safe_dump(clean, sort_keys=False, allow_unicode=True),
            )
            binary = _asset_binary(yaml_path)
            if binary is not None:
                ext = binary.suffix.lstrip(".") or "bin"
                zf.write(binary, f"assets/{aid}.{ext}")

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

        # index.json last so it reflects the actually-included content.
        index = _build_v5_index(skeletons, assets, themes)
        zf.writestr("index.json", json_mod.dumps(index, indent=2, ensure_ascii=False))

    blob = buf.getvalue()
    # Move staged → bundle so compose-run can later resolve the same
    # user asset ids the agent references in its plan.
    if user_meta:
        _move_staged_to_bundle()
    return blob


def _flat_prompt_text(skeletons: list[dict], assets: list[dict], brief: str) -> str:
    """Flat-text mirror of the v5 brief bundle. Includes SKILL_v5.md
    text, the v5 index.json, brand.md, and brief.md — i.e. the parts
    the agent could read as text. Binaries (themes/master.pptx,
    skeleton previews, asset images, user-asset previews) are not
    included; users wanting those should download the .zip.
    """
    # Pick themes the same way the zip path does so the flat index
    # matches the zip's contents.
    skeletons_root = cli_mod.WORKSPACE / "skeletons"  # noqa: F841 — kept for symmetry
    themes_root = cli_mod.WORKSPACE / "themes"
    decks_in_bundle = {sk.get("source_deck") for sk in skeletons}
    decks_in_bundle.discard(None)
    themes: list[dict] = []
    if themes_root.exists():
        for d in sorted(themes_root.iterdir()):
            if not d.is_dir() or d.name not in decks_in_bundle:
                continue
            theme_yaml = d / "theme.yaml"
            if not theme_yaml.exists():
                continue
            t = _safe_read(theme_yaml)
            if t:
                themes.append(t)

    sections: list[str] = []
    brand = _read_brand().strip()
    if brand:
        sections.append("=== brand.md ===\n" + brand)
    sections.append(
        "=== SKILL.md ===\n"
        + (cli_mod.CONSUMER / "SKILL_v5.md").read_text(encoding="utf-8")
    )
    index = _build_v5_index(skeletons, assets, themes)
    sections.append(
        "=== index.json ===\n" + json_mod.dumps(index, indent=2, ensure_ascii=False)
    )
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


@app.get("/api/compose/options")
def api_compose_options():
    return jsonify(_collect_filter_options())


@app.post("/api/compose/preview")
def api_compose_preview():
    body = request.get_json(force=True) or {}
    filters = body.get("filters") or {}
    skeletons, assets = _filter_kb(filters)
    return jsonify({
        "skeletons": len(skeletons),
        "assets": len(assets),
        "skeleton_ids": [s["id"] for s in skeletons],
        "asset_ids": [a["id"] for a in assets],
    })


@app.post("/api/compose/bundle")
def api_compose_bundle():
    body = request.get_json(force=True) or {}
    filters = body.get("filters") or {}
    brief = body.get("brief") or ""
    skeletons, assets = _filter_kb(filters)
    if not skeletons and not assets:
        debug_event("warn", "bundle",
                    "bundle request rejected — filters match nothing")
        return jsonify({"error": "filters match nothing — broaden them"}), 400
    user_count = len(_read_user_meta(USER_STAGED_DIR))
    blob = _build_prompt_bundle_zip(skeletons, assets, brief)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    debug_event(
        "info", "bundle",
        f"v5 brief bundle built — {len(skeletons)} skeletons, "
        f"{len(assets)} assets, {user_count} user assets, "
        f"{len(blob) // 1024} KB",
        skeletons=len(skeletons), assets=len(assets),
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
    skeletons, assets = _filter_kb(filters)
    return jsonify({"text": _flat_prompt_text(skeletons, assets, brief)})


@app.get("/api/vocab")
def api_vocab():
    # Ship the controlled vocab + the editable workspace tag vocabulary
    # in one payload so the describe UI can render tag chips without a
    # second round-trip. Build a fresh asset dict so we don't mutate
    # the module-level VOCAB by adding `tags` to it on every request.
    payload = dict(cli_mod.VOCAB)
    asset_vocab = dict(cli_mod.VOCAB.get("asset", {}))
    asset_vocab["tags"] = cli_mod.load_tag_vocab()
    payload["asset"] = asset_vocab
    return jsonify(payload)


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
                                  "description": ad.get("description", "")})

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
            "tags": [],
            "description": f"User-supplied {entry.get('kind') or 'asset'} "
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

    with tempfile.TemporaryDirectory(prefix="pptx-compose-") as tmpdir:
        staging = Path(tmpdir) / "bundle"
        staging.mkdir(parents=True, exist_ok=True)

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




@app.get("/")
def index():
    return render_template_string(UNIFIED_HTML)


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
    .msg.err pre { white-space: pre-wrap; word-break: break-word;
                   max-height: 240px; overflow: auto; margin: 6px 0 0;
                   font-family: ui-monospace, monospace; font-size: 11px;
                   background: #fff5f5; padding: 6px 8px; border-radius: 3px;
                   border: 1px solid #f0c8c8; color: #5a0000; }
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
    <a href="/">← review</a>
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
      <div class="group-label">Skeletons</div>
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
        Paste the JSON array the LLM returns. Entries are v5-shaped
        (<code>{"skeleton_id": …, "slots": …}</code>) and run via
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
const skFields = ["categories"];
const astFields = ["kind", "tags"];
const state = { skeletons: {}, assets: {} };

function fieldLabel(f) {
  return ({
    categories: "categories",
    kind: "kind",
    tags: "tags",
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
  if (state.options.skeletons && state.options.skeletons.options) {
    skFields.forEach((f) => {
      const vals = state.options.skeletons.options[f] || [];
      if (vals.length) buildDimension(tplEl, "skeletons", f, vals);
    });
  }
  if (state.options.assets && state.options.assets.options) {
    astFields.forEach((f) => {
      const vals = state.options.assets.options[f] || [];
      if (vals.length) buildDimension(astEl, "assets", f, vals);
    });
  }
  const skTotal = (state.options.skeletons && state.options.skeletons.total) || 0;
  const astTotal = (state.options.assets && state.options.assets.total) || 0;
  document.getElementById("kbSummary").textContent =
    `KB: ${skTotal} skeletons · ${astTotal} assets`;
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
    body: JSON.stringify({ filters: { skeletons: state.skeletons, assets: state.assets } }),
  });
  const j = await r.json();
  document.getElementById("countRow").textContent =
    `matching: ${j.skeletons} skeleton(s), ${j.assets} asset(s)`;
}

function currentFilters() {
  return { skeletons: state.skeletons, assets: state.assets };
}

function showMsg(id, text, ok) {
  const el = document.getElementById(id);
  el.className = "msg " + (ok ? "ok" : "err");
  el.textContent = text;
}

function showComposeError(j) {
  const el = document.getElementById("composeMsg");
  el.className = "msg err";
  el.textContent = "";
  const head = document.createElement("div");
  // Pull the most useful line out of stderr: the last "SomethingError: ..." line
  // is almost always the actual exception.
  const stderr = (j && j.stderr) || "";
  const lines = stderr.split(/\r?\n/).filter(Boolean);
  const errLine = [...lines].reverse().find(l => /^[A-Z][A-Za-z_]*(Error|Exception): /.test(l.trim()));
  const mode = (j && j.mode) ? " [" + j.mode + "]" : "";
  head.textContent = (j && j.error ? j.error : "compose failed") + mode +
                     (errLine ? "  —  " + errLine : "");
  el.appendChild(head);
  if (stderr) {
    const pre = document.createElement("pre");
    // Show the last ~60 lines so the relevant traceback frames are visible
    // without flooding the panel. Full stderr is also dumped to console.
    pre.textContent = lines.slice(-60).join("\n");
    el.appendChild(pre);
    console.error("compose stderr:\n" + stderr);
    if (j && j.stdout) console.error("compose stdout:\n" + j.stdout);
  }
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
    showComposeError(j);
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
    # Skeleton review lives on / now (Skeletons tab). Keep a redirect so
    # bookmarked URLs and the old `Skeletons →` links still work.
    return redirect("/", code=302)


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


UNIFIED_HTML = r"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>pptx-skill — review</title>
  <style>
    * { box-sizing: border-box; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
           margin: 0; height: 100vh; display: flex; color: #222; }

    .sidebar { width: 300px; border-right: 1px solid #ddd; background: #fafafa;
               display: flex; flex-direction: column; }
    .list-scroll { flex: 1; min-height: 0; overflow-y: auto; }
    .list-scroll.hidden { display: none; }
    .tab-bar { display: flex; border-bottom: 1px solid #ddd; background: #fff; flex-shrink: 0; }
    .tab-bar button { flex: 1; padding: 9px 8px; border: none; background: none;
                      cursor: pointer; font-weight: 500; font-size: 13px;
                      border-bottom: 2px solid transparent; color: #555; }
    .tab-bar button.active { border-bottom-color: #0066cc; color: #0066cc; font-weight: 600; }
    .tab-bar button:hover { background: #f5f5f5; }
    .filter-row { padding: 6px 14px; font-size: 12px; color: #555;
                  display: flex; gap: 10px; align-items: center;
                  border-bottom: 1px solid #eee; flex-shrink: 0; background: #fafafa; }
    .filter-row label { cursor: pointer; }
    .ingest-row { padding: 6px 14px; border-bottom: 1px solid #eee;
                  background: #fafafa; display: flex; align-items: center;
                  gap: 6px; flex-shrink: 0; }
    .ingest-row button { font-size: 11px; padding: 3px 8px;
                          border: 1px solid #0066cc; background: white;
                          color: #0066cc; border-radius: 3px; cursor: pointer; }
    .ingest-row button:hover { background: #eef4ff; }
    .ingest-row button:disabled { opacity: 0.5; cursor: not-allowed; }
    .ingest-row .ingest-msg { font-size: 10px; color: #777; flex: 1; min-width: 0;
                               overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .ingest-row .ingest-msg.ok { color: #1b5e20; }
    .ingest-row .ingest-msg.err { color: #a00; white-space: normal; font-size: 10px; }
    /* Asset list rows — id pill plus status */
    #assets-list .asset-item { padding: 8px 14px; cursor: pointer; font-size: 12px;
                                display: flex; justify-content: space-between;
                                align-items: center; border-bottom: 1px solid #eee; gap: 8px; }
    #assets-list .asset-item:hover { background: #eef4ff; }
    #assets-list .asset-item.active { background: #d8e8ff; }
    #assets-list .asset-id { font-family: ui-monospace, monospace; font-size: 11px;
                              color: #333; overflow: hidden; text-overflow: ellipsis;
                              white-space: nowrap; min-width: 0; flex: 1; }
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
    .item-list .sb-checkbox { margin: 0; flex-shrink: 0; cursor: pointer;
                              width: 14px; height: 14px; }
    .bulk-bar { flex-shrink: 0; background: #fff; border-top: 1px solid #ddd;
                padding: 10px 12px; display: flex; flex-direction: column; gap: 6px;
                box-shadow: 0 -2px 8px rgba(0,0,0,0.05); font-size: 12px; }
    .bulk-bar.hidden { display: none; }
    .bulk-bar .count { color: #555; font-weight: 600; }
    .bulk-bar .row { display: flex; gap: 6px; flex-wrap: wrap; }
    .bulk-bar button { font-size: 11px; padding: 5px 10px; border: 1px solid #ccc;
                       background: #fff; border-radius: 4px; cursor: pointer; }
    .bulk-bar button.primary { background: #2e7d32; color: #fff; border-color: #2e7d32; }
    .bulk-bar button.danger { background: #c62828; color: #fff; border-color: #c62828; }
    .bulk-bar button.ghost { background: #f5f5f5; }
    .bulk-bar button:hover { filter: brightness(1.08); }

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
    .slot-overlay.hover-highlight { box-shadow: 0 0 0 3px #ffc107, 0 0 14px 2px rgba(255,193,7,0.55);
                                    opacity: 1; z-index: 50; }
    .slot-overlay.hover-highlight .slot-label { background: #ffc107; color: #000; }

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

    /* Asset describe form (Assets tab right panel) */
    #asset-form label { display: block; font-size: 11px; font-weight: 700;
                         margin: 12px 0 4px; color: #555; text-transform: uppercase;
                         letter-spacing: 0.4px; }
    #asset-form label .hint { font-weight: 400; color: #999; text-transform: none;
                               letter-spacing: 0; margin-left: 6px; font-size: 11px; }
    #asset-form input, #asset-form select, #asset-form textarea {
      width: 100%; padding: 6px 8px; border: 1px solid #ccc; border-radius: 4px;
      font-size: 13px; font-family: inherit; background: white;
    }
    #asset-form textarea { resize: vertical; min-height: 50px; }
    #asset-form .checks { display: flex; flex-wrap: wrap; gap: 4px; }
    #asset-form .checks label { display: inline-flex; align-items: center; gap: 4px;
                                 font-weight: normal; font-size: 11px; padding: 3px 8px;
                                 border: 1px solid #ccc; border-radius: 12px; cursor: pointer;
                                 background: #fff; margin: 0; text-transform: none;
                                 letter-spacing: 0; }
    #asset-form .checks input { width: auto; margin: 0; }
    #asset-form .checks label:has(input:checked) { background: #d8e8ff; border-color: #0066cc;
                                                    color: #003e7e; font-weight: 600; }
    .asset-btnrow { display: flex; gap: 8px; margin-top: 16px; flex-wrap: wrap; }
    .asset-btnrow button { font-size: 12px; padding: 6px 12px; border-radius: 4px;
                           border: 1px solid #ccc; background: white; cursor: pointer; }
    .asset-btnrow button.primary { background: #0066cc; color: white; border-color: #0066cc; }
    .asset-btnrow button.primary:hover { background: #0052a3; }
    .asset-btnrow button:hover { background: #f0f0f0; }
    #asset-msg .errors { background: #ffe9e9; color: #a00; padding: 8px;
                          border-radius: 4px; margin-top: 10px; font-size: 12px;
                          border: 1px solid #f5b8b8; }
    #asset-msg .errors ul { margin: 4px 0 0 16px; padding: 0; }
    #asset-msg .ok { color: #1b5e20; font-size: 12px; margin-top: 8px;
                      background: #e8f5e9; padding: 8px; border-radius: 4px;
                      border: 1px solid #c8e6c9; }
    .asset-preview-wrap img { max-width: 100%; max-height: calc(100vh - 80px);
                              object-fit: contain; background: white;
                              box-shadow: 0 4px 30px rgba(0,0,0,0.4); }

    /* Belt-and-suspenders: hide v4 debug-floater if it leaks in. */
    .dbg-floater { display: none !important; }
  </style>
</head>
<body>
  <aside class="sidebar">
    <header>
      <h1>pptx-skill — review</h1>
      <div class="nav"><a href="/compose">compose →</a></div>
    </header>
    <div class="tab-bar" id="tab-bar">
      <button data-tab="skeletons" class="active">Skeletons</button>
      <button data-tab="assets">Assets</button>
    </div>
    <div class="ingest-row" id="ingest-row-skeletons">
      <input type="file" id="ingestFile" accept=".pptx" style="display:none;" />
      <button id="ingestBtn" type="button">+ Ingest .pptx</button>
      <span class="ingest-msg" id="ingestMsg"
            title="Upload a .pptx to extract skeletons + assets. Re-ingesting an existing deck is rejected — delete its workspace/decks/ dir first.">
        Add a new deck
      </span>
    </div>
    <div class="ingest-row hidden" id="ingest-row-assets">
      <input type="file" id="addAssetFile"
             accept=".png,.jpg,.jpeg,.webp,.gif,.svg,.xml" style="display:none;" />
      <button id="addAssetBtn" type="button">+ Add asset</button>
      <span class="ingest-msg" id="addAssetMsg"
            title="Upload a single image / SVG / XML — lands as a pending asset, ready to describe.">
        Register a single image
      </span>
    </div>
    <div class="filter-row hidden" id="filter-row-assets">
      <label><input type="checkbox" id="hideDoneAssets"> Hide done</label>
    </div>
    <div id="skeletons-list" class="list-scroll"></div>
    <div id="assets-list" class="list-scroll hidden"></div>
    <div id="bulk-bar" class="bulk-bar hidden"></div>
  </aside>
  <main class="preview-area" id="preview-area">
    <div class="preview-empty">
      <div>Select a skeleton or asset on the left.</div>
      <div class="hint">tab switcher above the list · skeletons = layouts · assets = images</div>
    </div>
  </main>
  <aside class="panel" id="panel">
    <div class="preview-empty" style="text-align:center;padding-top:40px;">
      Nothing selected.
    </div>
  </aside>

  <script>
    const KIND_LIST = ['heading', 'paragraph', 'bullets', 'table', 'chart', 'image', 'footer'];
    const ROLE_LIST = ['(none)', 'page_title', 'subtitle', 'body', 'footer', 'date',
      'page_number', 'footnote', 'caption', 'key_points', 'detailed_list',
      'cta', 'byline', 'kpi_label', 'kpi_value', 'section_header'];

    // Multi-select state survives across loadSkeletons() reloads so a
    // bulk action followed by the list re-render doesn't drop selection.
    // We prune stale ids after each reload.
    const selected = new Set();

    async function loadSkeletons() {
      const r = await fetch('/api/v5/skeletons');
      const data = await r.json();
      const root = document.getElementById('skeletons-list');
      root.innerHTML = '';
      const decks = Object.keys(data.decks).sort();
      if (decks.length === 0) {
        root.innerHTML = '<div style="padding:14px;color:#888;font-size:12px;line-height:1.5;">No skeletons yet.<br>Run:<br><code style="font-size:11px;">python3 authoring/cli.py ingest your_deck.pptx</code></div>';
        selected.clear();
        renderBulkBar();
        return;
      }
      const allIds = new Set();
      decks.forEach(deck => {
        const group = document.createElement('div');
        group.className = 'deck-group';
        const h = document.createElement('h2');
        h.textContent = deck;
        group.appendChild(h);
        const ul = document.createElement('ul');
        ul.className = 'item-list';
        data.decks[deck].forEach(sk => {
          allIds.add(sk.id);
          const li = document.createElement('li');
          li.dataset.id = sk.id;
          const cats = (sk.categories || []).join(', ') || '—';
          const warn = sk.has_warnings ? '<span class="warn-dot" title="overlap_detected">⚠</span> ' : '';
          const checked = selected.has(sk.id) ? 'checked' : '';
          li.innerHTML = `
            <input type="checkbox" class="sb-checkbox" ${checked} title="select for bulk action">
            <span class="label">
              <span class="top">${warn}slide ${sk.source_slide_index} · ${sk.slot_count} slots</span>
              <span class="cats">${cats}</span>
            </span>
            <span class="pill ${sk.status}">${sk.status}</span>
          `;
          const cb = li.querySelector('.sb-checkbox');
          // Stop click on checkbox from bubbling to <li> and triggering selectSkeleton.
          cb.addEventListener('click', e => e.stopPropagation());
          cb.addEventListener('change', e => toggleSelect(sk.id, e.target.checked));
          li.onclick = () => selectSkeleton(sk.id);
          ul.appendChild(li);
        });
        group.appendChild(ul);
        root.appendChild(group);
      });
      // Drop any selected ids that no longer exist (e.g. after a delete).
      for (const id of Array.from(selected)) {
        if (!allIds.has(id)) selected.delete(id);
      }
      renderBulkBar();
    }

    function toggleSelect(id, on) {
      if (on) selected.add(id); else selected.delete(id);
      renderBulkBar();
    }

    function clearSelection() {
      selected.clear();
      document.querySelectorAll('.sb-checkbox').forEach(cb => { cb.checked = false; });
      renderBulkBar();
    }

    function renderBulkBar() {
      const bar = document.getElementById('bulk-bar');
      if (!bar) return;
      if (selected.size === 0) {
        bar.classList.add('hidden');
        bar.innerHTML = '';
        return;
      }
      bar.classList.remove('hidden');
      bar.innerHTML = `
        <div class="count">${selected.size} selected</div>
        <div class="row">
          <button class="primary" onclick="bulkSetStatus('done')">Mark done</button>
          <button class="danger" onclick="bulkSetStatus('rejected')">Reject</button>
          <button class="ghost" onclick="clearSelection()">Clear</button>
        </div>
      `;
    }

    async function bulkSetStatus(status) {
      if (selected.size === 0) return;
      const ids = Array.from(selected);
      // Big bulk actions are easy to fat-finger; ask once.
      if (ids.length >= 5) {
        const verb = status === 'done' ? 'mark as done' : 'reject';
        if (!confirm(`${verb} ${ids.length} skeletons?`)) return;
      }
      const results = await Promise.all(ids.map(id =>
        fetch(`/api/v5/skeleton/${encodeURIComponent(id)}/status`, {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify({status})
        }).then(r => r.ok).catch(() => false)
      ));
      const failed = ids.filter((_, i) => !results[i]);
      if (failed.length) {
        alert(`Failed for ${failed.length} of ${ids.length}: ${failed.join(', ')}`);
        // Keep failed ids selected so the user can see + retry.
        selected.clear();
        failed.forEach(id => selected.add(id));
      } else {
        selected.clear();
      }
      await loadSkeletons();
    }

    // Walk the rendered sidebar in DOM order (= deck × source_slide order)
    // and return the next still-pending id strictly AFTER currentId, or
    // null. Forward-only so an intentionally-left-pending earlier slide
    // isn't surprise-loaded after the user approves a later one.
    function findNextPending(currentId) {
      const lis = Array.from(document.querySelectorAll('.item-list li'));
      const items = lis.map(li => ({
        id: li.dataset.id,
        pending: !!li.querySelector('.pill.pending'),
      }));
      const startIdx = items.findIndex(i => i.id === currentId);
      if (startIdx === -1) return null;
      for (let i = startIdx + 1; i < items.length; i++) {
        if (items[i].pending) return items[i].id;
      }
      return null;
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
      await loadSkeletons();
      // Auto-advance only when moving AWAY from pending. Corrective
      // actions (back-to-pending, restore) should keep the user on the
      // current skeleton so they can keep editing it.
      if (status === 'done' || status === 'rejected') {
        const nextId = findNextPending(id);
        if (nextId) { await selectSkeleton(nextId); return; }
      }
      await refreshCurrent(id);
    }

    async function setOverlapDecision(id, decision) {
      const r = await fetch(`/api/v5/skeleton/${encodeURIComponent(id)}/overlap-decision`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({decision})
      });
      if (!r.ok) { alert('Overlap decision failed: ' + (await r.text())); return; }
      await loadSkeletons();
      // 'reject' decision flips status to rejected — treat like setStatus(reject).
      if (decision === 'reject') {
        const nextId = findNextPending(id);
        if (nextId) { await selectSkeleton(nextId); return; }
      }
      await refreshCurrent(id);
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
      document.querySelectorAll(`.slot-overlay.unmapped[data-unmapped-index="${idx}"]`).forEach(o => {
        o.classList.toggle('hover-highlight', on);
      });
    }

    function highlightSlot(slotId, on) {
      document.querySelectorAll(`.slot-card[data-slot-id="${CSS.escape(slotId)}"]`).forEach(c => {
        c.classList.toggle('active', on);
      });
      document.querySelectorAll(`.slot-overlay[data-slot-id="${CSS.escape(slotId)}"]`).forEach(o => {
        o.classList.toggle('hover-highlight', on);
      });
    }

    // Reverse highlight: hover over a panel card -> glow the preview overlay.
    // Event delegation on #panel so it survives innerHTML re-renders.
    (function wirePanelHoverHighlight() {
      const panel = document.getElementById('panel');
      if (!panel || panel.dataset.hoverWired) return;
      panel.dataset.hoverWired = '1';
      panel.addEventListener('mouseover', e => {
        const slotCard = e.target.closest('.slot-card[data-slot-id]');
        if (slotCard && !slotCard.contains(e.relatedTarget)) {
          highlightSlot(slotCard.dataset.slotId, true); return;
        }
        const unCard = e.target.closest('.unmapped-card[data-unmapped-index]');
        if (unCard && !unCard.contains(e.relatedTarget)) {
          highlightUnmapped(Number(unCard.dataset.unmappedIndex), true);
        }
      });
      panel.addEventListener('mouseout', e => {
        const slotCard = e.target.closest('.slot-card[data-slot-id]');
        if (slotCard && !slotCard.contains(e.relatedTarget)) {
          highlightSlot(slotCard.dataset.slotId, false); return;
        }
        const unCard = e.target.closest('.unmapped-card[data-unmapped-index]');
        if (unCard && !unCard.contains(e.relatedTarget)) {
          highlightUnmapped(Number(unCard.dataset.unmappedIndex), false);
        }
      });
    })();

    function escapeHtml(s) {
      return (s == null ? '' : String(s)).replace(/[&<>"']/g,
        c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    }

    // ===========================================================================
    // Tab switching: skeletons <-> assets
    // ===========================================================================

    let activeTab = localStorage.getItem('reviewTab') || 'skeletons';
    let ASSET_KIND = [];
    let ASSET_TAGS = [];
    let assetItems = [];        // array of {id, yaml, status}
    let currentAsset = null;    // {kind, yaml, data, yaml_text} from /api/item

    function setTab(tab) {
      activeTab = tab;
      localStorage.setItem('reviewTab', tab);
      document.querySelectorAll('.tab-bar button').forEach(b =>
        b.classList.toggle('active', b.dataset.tab === tab));
      const showSk = tab === 'skeletons';
      document.getElementById('skeletons-list').classList.toggle('hidden', !showSk);
      document.getElementById('assets-list').classList.toggle('hidden', showSk);
      document.getElementById('ingest-row-skeletons').classList.toggle('hidden', !showSk);
      document.getElementById('ingest-row-assets').classList.toggle('hidden', showSk);
      document.getElementById('filter-row-assets').classList.toggle('hidden', showSk);
      // Multi-select is skeleton-only — clear any held selection when leaving.
      if (!showSk && selected.size) clearSelection();
      // Reset main panes — switching tab should not show stale detail.
      resetMainPanes();
      if (tab === 'assets') {
        if (!ASSET_KIND.length) loadAssetVocab();
        loadAssets();
      } else {
        loadSkeletons();
      }
    }

    function resetMainPanes() {
      const area = document.getElementById('preview-area');
      area.innerHTML = `<div class="preview-empty"><div>Select an item on the left.</div></div>`;
      const panel = document.getElementById('panel');
      panel.innerHTML = `<div class="preview-empty" style="text-align:center;padding-top:40px;">Nothing selected.</div>`;
    }

    document.querySelectorAll('.tab-bar button').forEach(b => {
      b.addEventListener('click', () => setTab(b.dataset.tab));
    });

    // ===========================================================================
    // Assets tab — list + select + describe form
    // ===========================================================================

    async function loadAssetVocab() {
      const r = await fetch('/api/vocab');
      if (!r.ok) return;
      const v = await r.json();
      ASSET_KIND = (v.asset && v.asset.kind) || [];
      ASSET_TAGS = (v.asset && v.asset.tags) || [];
    }

    async function loadAssets() {
      const r = await fetch('/api/items');
      const data = await r.json();
      assetItems = data.assets || [];
      renderAssetList();
    }

    function renderAssetList() {
      const root = document.getElementById('assets-list');
      const hideDone = document.getElementById('hideDoneAssets').checked;
      const visible = hideDone ? assetItems.filter(a => a.status !== 'done') : assetItems;
      if (!visible.length) {
        root.innerHTML = `<div style="padding:14px;color:#888;font-size:12px;">
          No assets${hideDone ? ' pending' : ''}.</div>`;
        return;
      }
      root.innerHTML = '';
      visible.forEach(it => {
        const div = document.createElement('div');
        div.className = 'asset-item';
        div.dataset.yaml = it.yaml;
        if (currentAsset && currentAsset.yaml === it.yaml) div.classList.add('active');
        div.innerHTML = `<span class="asset-id" title="${it.yaml}">${escapeHtml(it.id)}</span>
          <span class="pill ${it.status}">${it.status}</span>`;
        div.addEventListener('click', () => selectAsset(it.yaml));
        root.appendChild(div);
      });
    }

    document.getElementById('hideDoneAssets').addEventListener('change', renderAssetList);

    async function selectAsset(yamlRel) {
      const r = await fetch('/api/item?yaml=' + encodeURIComponent(yamlRel));
      const item = await r.json();
      currentAsset = item;
      renderAssetPreview(yamlRel);
      renderAssetPanel(item);
      renderAssetList();
    }

    function renderAssetPreview(yamlRel) {
      const area = document.getElementById('preview-area');
      area.innerHTML = `<div class="asset-preview-wrap">
        <img src="/preview?yaml=${encodeURIComponent(yamlRel)}&t=${Date.now()}" alt="preview">
      </div>`;
    }

    function renderAssetPanel(item) {
      const d = item.data || {};
      const panel = document.getElementById('panel');
      panel.innerHTML = `
        <h2>${escapeHtml(d.id || item.yaml)}</h2>
        <div class="subtitle">${escapeHtml(item.yaml)}</div>
        <form id="asset-form" onsubmit="event.preventDefault(); saveAsset(true);"></form>
        <div class="asset-btnrow">
          <button type="button" class="primary" onclick="saveAsset(true)">Save + next</button>
          <button type="button" onclick="saveAsset(false)">Save</button>
          <button type="button" onclick="copyAssetPrompt()" id="copyAssetPromptBtn">Copy describer prompt</button>
        </div>
        <div id="asset-msg"></div>
      `;
      const f = document.getElementById('asset-form');
      addAssetSelect(f, 'kind', d.kind || '', ASSET_KIND);
      addAssetChips(f, 'tags', d.tags || [], ASSET_TAGS);
      addAssetText(f, 'description', d.description || '',
        'one short sentence, <=25 words, what is literally visible');
      addAssetTextarea(f, 'notes', d.notes || '', 'human reviewer note');
    }

    function addAssetText(parent, name, value, hint) {
      parent.insertAdjacentHTML('beforeend',
        `<label for="af-${name}">${name}${hint ? ' <span class="hint">' + hint + '</span>' : ''}</label>
         <input type="text" id="af-${name}" name="${name}" value="${escapeHtml(value)}">`);
    }
    function addAssetTextarea(parent, name, value, hint) {
      parent.insertAdjacentHTML('beforeend',
        `<label for="af-${name}">${name}${hint ? ' <span class="hint">' + hint + '</span>' : ''}</label>
         <textarea id="af-${name}" name="${name}" rows="3">${escapeHtml(value)}</textarea>`);
    }
    function addAssetSelect(parent, name, value, options) {
      const opts = ['<option value="">—</option>'].concat(
        options.map(o => `<option value="${escapeHtml(o)}"${o === value ? ' selected' : ''}>${escapeHtml(o)}</option>`)
      ).join('');
      parent.insertAdjacentHTML('beforeend',
        `<label for="af-${name}">${name}</label>
         <select id="af-${name}" name="${name}">${opts}</select>`);
    }
    function addAssetChips(parent, name, values, options) {
      const valSet = new Set(values);
      const chips = options.map(o =>
        `<label><input type="checkbox" value="${escapeHtml(o)}"${valSet.has(o) ? ' checked' : ''}> ${escapeHtml(o)}</label>`
      ).join('');
      parent.insertAdjacentHTML('beforeend',
        `<label>${name}</label>
         <div class="checks" id="af-${name}">${chips || '<span style="color:#888;font-size:11px;">(no tags in vocab; add via cli.py tag-vocab add)</span>'}</div>`);
    }

    function gatherAssetForm() {
      const tags = [...document.querySelectorAll('#af-tags input:checked')].map(c => c.value);
      return {
        kind: document.getElementById('af-kind').value,
        tags,
        description: document.getElementById('af-description').value.trim(),
        notes: document.getElementById('af-notes').value.trim(),
      };
    }

    async function saveAsset(advance) {
      if (!currentAsset) return;
      const fields = gatherAssetForm();
      const r = await fetch('/api/save', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({yaml: currentAsset.yaml, fields}),
      });
      const result = await r.json();
      const msg = document.getElementById('asset-msg');
      if (result.errors && result.errors.length) {
        msg.innerHTML = '<div class="errors"><strong>Validation errors</strong><ul>'
          + result.errors.map(e => '<li>' + escapeHtml(e) + '</li>').join('')
          + '</ul></div>';
        return;
      }
      msg.innerHTML = '<div class="ok">Saved — status: ' + escapeHtml(result.status || '') + '</div>';
      const it = assetItems.find(i => i.yaml === currentAsset.yaml);
      if (it) it.status = result.status;
      renderAssetList();
      if (advance) {
        const next = assetItems.find(i => i.status === 'pending' && i.yaml !== currentAsset.yaml);
        if (next) { await selectAsset(next.yaml); }
        else { msg.innerHTML += '<div class="ok" style="margin-top:6px;">No more pending assets 🎉</div>'; }
      }
    }

    async function copyAssetPrompt() {
      const r = await fetch('/api/prompt');
      const j = await r.json();
      try {
        await navigator.clipboard.writeText(j.text);
        const btn = document.getElementById('copyAssetPromptBtn');
        const orig = btn.textContent;
        btn.textContent = 'Copied ✓';
        setTimeout(() => (btn.textContent = orig), 1500);
      } catch (e) {
        alert('Copy failed:\n' + e.message);
      }
    }

    // ===========================================================================
    // Ingest + add-asset uploads (existing endpoints, unchanged behavior)
    // ===========================================================================

    document.getElementById('ingestBtn').addEventListener('click', () => {
      document.getElementById('ingestFile').click();
    });
    document.getElementById('ingestFile').addEventListener('change', async (ev) => {
      const file = ev.target.files[0];
      if (!file) return;
      const msg = document.getElementById('ingestMsg');
      msg.textContent = `Ingesting ${file.name}…`;
      msg.className = 'ingest-msg';
      const fd = new FormData(); fd.append('pptx', file);
      try {
        const r = await fetch('/api/ingest', {method: 'POST', body: fd});
        const j = await r.json();
        if (!r.ok) {
          msg.textContent = j.error || 'ingest failed';
          msg.className = 'ingest-msg err';
        } else {
          msg.textContent = `Ingested ${j.deck_stem}: ${j.slides} slides, ${j.pictures + j.atoms} assets`;
          msg.className = 'ingest-msg ok';
          if (activeTab === 'skeletons') loadSkeletons(); else loadAssets();
        }
      } catch (e) {
        msg.textContent = 'ingest failed: ' + e.message;
        msg.className = 'ingest-msg err';
      } finally {
        ev.target.value = '';
      }
    });

    document.getElementById('addAssetBtn').addEventListener('click', () => {
      document.getElementById('addAssetFile').click();
    });
    document.getElementById('addAssetFile').addEventListener('change', async (ev) => {
      const file = ev.target.files[0];
      if (!file) return;
      const msg = document.getElementById('addAssetMsg');
      msg.textContent = `Uploading ${file.name}…`;
      msg.className = 'ingest-msg';
      const fd = new FormData(); fd.append('file', file);
      try {
        const r = await fetch('/api/asset/add', {method: 'POST', body: fd});
        const j = await r.json();
        if (!r.ok) {
          msg.textContent = j.error || 'add-asset failed';
          msg.className = 'ingest-msg err';
        } else {
          msg.textContent = `Added ${j.id} (${j.kind})`;
          msg.className = 'ingest-msg ok';
          if (activeTab === 'assets') loadAssets();
        }
      } catch (e) {
        msg.textContent = 'add-asset failed: ' + e.message;
        msg.className = 'ingest-msg err';
      } finally {
        ev.target.value = '';
      }
    });

    // ===========================================================================
    // Boot
    // ===========================================================================

    (async () => {
      await loadAssetVocab();
      setTab(activeTab);  // calls loadSkeletons() or loadAssets() as appropriate
    })();
  </script>
</body>
</html>
"""


def main():
    port = int(os.environ.get("PPTX_SKILL_PORT", "5050"))
    # `/` is the unified review page (Skeletons | Assets tabs).
    # /v5 redirects to / for old bookmarks.
    landing = os.environ.get("PPTX_SKILL_LANDING", "/")
    url = f"http://127.0.0.1:{port}{landing}"
    if not os.environ.get("PPTX_SKILL_NO_BROWSER"):
        Timer(1.2, lambda: webbrowser.open(url)).start()
    print(f"pptx-skill app → {url}")
    print(f"  review:    http://127.0.0.1:{port}/")
    print(f"  compose:   http://127.0.0.1:{port}/compose")
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
