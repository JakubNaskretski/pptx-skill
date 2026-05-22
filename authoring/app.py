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
import shutil
import subprocess
import sys
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

app = Flask(__name__)


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
    png = slide_pptx.with_suffix(".png")
    if png.exists() and png.stat().st_mtime >= slide_pptx.stat().st_mtime:
        return png
    ql = shutil.which("qlmanage")
    if not ql:
        return None
    try:
        subprocess.run(
            [ql, "-t", "-s", "1200", "-o", str(slide_pptx.parent), str(slide_pptx)],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=20,
        )
    except (subprocess.SubprocessError, OSError):
        return None
    # qlmanage writes <stem>.pptx.png — rename to <stem>.png so it matches
    # the convention used by `cli.py preview` and `cli.py build`.
    produced = slide_pptx.parent / f"{slide_pptx.name}.png"
    if produced.exists():
        produced.replace(png)
    return png if png.exists() else None


def _asset_binary(yaml_path: Path) -> Path | None:
    for cand in yaml_path.parent.glob(f"{yaml_path.stem}.*"):
        if cand.suffix != ".yaml":
            return cand
    return None


SLIDE_DESCRIPTIVE = ("intent", "feel", "suitable_for", "notes")
ASSET_DESCRIPTIVE = (
    "kind", "subject", "depicts", "feel", "composition",
    "colors", "scope", "suitable_for", "notes",
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
    txt = txt.strip()
    if txt.startswith("```yaml"):
        txt = txt[len("```yaml"):].lstrip()
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


def _bulk_instructions(kind: str, n: int, per_item_prompt: str) -> str:
    item_name = "image" if kind == "asset" else "slide preview"
    if kind == "asset":
        sample_block = (
            '"01":\n'
            '  kind: photo\n'
            '  subject: "..."\n'
            '  depicts: "..."\n'
            '  feel: warm\n'
            '  composition: centered\n'
            '  colors: [navy, white]\n'
            '  scope: [generic]\n'
            '  suitable_for: [team]\n'
            '  notes: ""\n'
            '\n'
            '"02":\n'
            '  kind: photo\n'
            '  # ... same fields ...\n'
        )
    else:
        sample_block = (
            '"01":\n'
            '  intent: "..."\n'
            '  feel: formal\n'
            '  suitable_for: [opener]\n'
            '  notes: ""\n'
            '\n'
            '"02":\n'
            '  intent: "..."\n'
            '  # ... same fields ...\n'
        )
    return (
        f"# Bulk describe batch — {n} {item_name}s\n\n"
        f"You will see {n} {item_name}s numbered 01 through {n:02d}. For each, "
        f"produce a description following the schema in the second half of "
        f"this file.\n\n"
        f"## CRITICAL — Output format rules\n\n"
        f"Return ONE YAML mapping. The top-level keys MUST be the quoted "
        f"2-digit ids: `\"01\"`, `\"02\"`, ..., `\"{n:02d}\"`.\n\n"
        f"**Every per-item field MUST be indented by exactly 2 spaces under "
        f"its id.** Failure to indent breaks the parser and the response "
        f"will be unusable.\n\n"
        f"### CORRECT (note the 2-space indent before each field):\n\n"
        f"```yaml\n{sample_block}```\n\n"
        f"### INCORRECT — do NOT do this:\n\n"
        f"```yaml\n"
        f'"01":\n'
        f'intent: "..."     # WRONG — must be indented under "01"\n'
        f'feel: formal      # WRONG\n'
        f'"02":\n'
        f'intent: "..."     # WRONG\n'
        f"```\n\n"
        f"Return EXACTLY {n} entries, one per item. Do NOT skip any. Output "
        f"ONLY the YAML mapping. No commentary, no markdown code fences, no "
        f"prose before or after.\n\n"
        f"---\n\n"
        f"## Per-item description schema\n\n"
        f"{per_item_prompt}\n"
    )


_FLAT_ID_RE = None


def _recover_flat_batch_yaml(raw_text: str) -> str:
    """LLMs often emit per-item fields at column 0 instead of indented under
    each id key. Detect that pattern and re-indent before parsing.

    Returns either the original text (no fixup needed) or a re-indented
    version that PyYAML can parse correctly.
    """
    import re
    global _FLAT_ID_RE
    if _FLAT_ID_RE is None:
        _FLAT_ID_RE = re.compile(r'^"?(\d{1,4})"?\s*:\s*$')
    lines = raw_text.split("\n")
    id_lines = [i for i, ln in enumerate(lines) if _FLAT_ID_RE.match(ln)]
    if len(id_lines) < 2:
        return raw_text
    # Need recovery if any non-blank, non-id line between id lines starts at col 0
    needs = False
    in_block = False
    for ln in lines:
        if _FLAT_ID_RE.match(ln):
            in_block = True
            continue
        if not in_block or ln.strip() == "":
            continue
        if not ln.startswith((" ", "\t")):
            needs = True
            break
    if not needs:
        return raw_text
    out = []
    in_block = False
    for ln in lines:
        if _FLAT_ID_RE.match(ln):
            in_block = True
            out.append(ln)
            continue
        if not in_block:
            out.append(ln)
            continue
        if ln.strip() == "" or ln.startswith((" ", "\t")):
            out.append(ln)
            continue
        out.append("  " + ln)
    return "\n".join(out)


@app.post("/api/batch/create")
def api_batch_create():
    body = request.get_json(force=True) or {}
    kind = body.get("kind", "asset")
    try:
        count = max(1, min(20, int(body.get("count", 10))))
    except (TypeError, ValueError):
        return jsonify({"error": "count must be an integer"}), 400
    if kind not in ("asset", "slide"):
        return jsonify({"error": "kind must be 'asset' or 'slide'"}), 400

    if kind == "asset":
        candidates = list(cli_mod.iter_asset_yamls())
    else:
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
                else:
                    slide_pptx = ypath.with_suffix(".pptx")
                    if not slide_pptx.exists():
                        skipped.append({"yaml": rel, "reason": "slide fragment .pptx missing"})
                        continue
                    png = _ensure_slide_png(slide_pptx)
                    if png is None:
                        if not shutil.which("qlmanage"):
                            reason = "qlmanage not available (install macOS or run preview manually)"
                        else:
                            reason = "qlmanage failed to render this slide"
                        skipped.append({"yaml": rel, "reason": reason})
                        continue
                    key = f"{added + 1:02d}"
                    zf.write(png, f"{key}.png")
                manifest["items"][key] = rel
                added += 1
            except OSError as e:
                skipped.append({"yaml": rel, "reason": f"OS error: {e}"})
            except Exception as e:
                skipped.append({"yaml": rel, "reason": f"unexpected: {type(e).__name__}: {e}"})

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

    return jsonify({
        "batch_id": batch_id,
        "kind": kind,
        "count": added,
        "requested": count,
        "items": manifest["items"],
        "skipped": skipped,
        "download_url": f"/api/batch/{batch_id}/download",
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
    txt = _recover_flat_batch_yaml(txt)
    try:
        parsed = yaml.safe_load(txt) if txt else None
    except Exception as e:
        return jsonify({"error": f"YAML parse error: {e}"}), 400
    if not isinstance(parsed, (dict, list)):
        return jsonify({"error": "expected a YAML mapping or list at top level"}), 400

    items_dict = _find_items_dict(parsed) or {}
    found_keys = [str(k) for k in items_dict.keys()]
    by_int: dict[int, dict] = {}
    for k, v in items_dict.items():
        n = _normalize_key_to_int(k)
        if n is not None and isinstance(v, dict):
            by_int[n] = v

    results = []
    for key, rel in manifest["items"].items():
        n = _normalize_key_to_int(key)
        entry = by_int.get(n) if n is not None else None
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
        "matched": sum(1 for r in results if r["status"] != "no-match"),
    })


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
    return jsonify({"batches": out[:20]})


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
  </style>
</head>
<body>
  <aside class="sidebar">
    <header>
      <strong>pptx-skill</strong>
      <span id="counts" style="color:#888;font-size:11px;margin-left:6px;"></span>
    </header>
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
        LLM. The model returns one YAML mapping; paste it back below to apply
        all descriptions at once.
      </p>
      <div class="inline">
        <label>Kind:
          <select id="batchKind">
            <option value="asset">Assets</option>
            <option value="slide">Slides</option>
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
        Paste the YAML mapping the LLM returned. Items are matched by id
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
      <textarea id="batchYaml" placeholder='"01":&#10;  kind: photo&#10;  subject: ...&#10;&#10;"02":&#10;  ...'></textarea>
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
const SLIDE_FEEL = ["formal","punchy","data-dense","warm","clinical","celebratory"];
const SLIDE_TAGS = ["opener","section_divider","content","data","quote","closing","product","team"];
const ASSET_KIND = ["photo","icon","logo","illustration","screenshot"];
const ASSET_FEEL = ["formal","warm","clinical","punchy","playful","minimal","dramatic"];
const ASSET_COMP = ["centered","left-weighted","right-weighted","full-bleed","top-heavy","scattered"];
const ASSET_TAGS = ["team","hero","product","data","culture","event","abstract","decorative","closing","quote"];

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
    addTextarea(f, "notes", d.notes || "");
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
    addTextarea(f, "notes", d.notes || "");
  }
  document.getElementById("msg").innerHTML = "";
}

function addText(parent, name, value, hint) {
  const lbl = el("label", {for: name}, name);
  if (hint) lbl.append(el("span", {class: "hint"}, hint));
  parent.append(lbl);
  parent.append(el("input", {type: "text", name, value, id: name}));
}
function addTextarea(parent, name, value) {
  parent.append(el("label", {for: name}, name));
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
  document.getElementById("preview").hidden = !inItems;
  document.getElementById("panel").hidden = !inItems;
  document.getElementById("batchView").hidden = inItems;
  if (!inItems) loadBatches();
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
  hint.textContent = `(${n} pending ${kind}${n === 1 ? "" : "s"})`;
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

  const itemsList = Object.entries(j.items).map(([k, v]) =>
    `<code>${k}</code> → ${escapeHtml(_shortPath(v))}`
  ).join("<br>");

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
  window.location.href = j.download_url;
}

async function batchApply() {
  if (!currentBatchId) {
    alert("Generate a batch first.");
    return;
  }
  const text = document.getElementById("batchYaml").value;
  if (!text.trim()) { alert("Paste the LLM's YAML first."); return; }
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
});
document.getElementById("batchRefreshBtn").addEventListener("click", () => loadBatches());
document.getElementById("hideDone").addEventListener("change", renderList);
document.getElementById("saveBtn").addEventListener("click", () => save(true));
document.getElementById("saveOnly").addEventListener("click", () => save(false));
document.getElementById("copyPrompt").addEventListener("click", copyPrompt);
document.querySelectorAll(".mode-toggle button").forEach(b => {
  b.addEventListener("click", () => setMode(b.dataset.mode));
});

applyMode();
applyView();
loadItems();
</script>
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(INDEX_HTML)


def main():
    url = "http://127.0.0.1:5000/"
    Timer(1.2, lambda: webbrowser.open(url)).start()
    print(f"pptx-skill describe app → {url}")
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
