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