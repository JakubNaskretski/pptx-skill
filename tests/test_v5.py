"""Tests for pptx-skill v5 — structural-skeleton pipeline.

Run from repo root: ``python3 -m unittest tests.test_v5``

unittest-only (no pytest dep). Synthetic decks built in-memory via
python-pptx; no binary fixtures committed.

Covers:
- ingest_v5.digest_skeleton — slot inventory + unmapped_shapes + shape_id
- ingest_v5.compute_repeated_picture_info — median area + threshold
- ingest_v5._propose_categories — opening / data / closing rules
- reader.cmd_v5_match_skeletons — ranking + zero-match suggested_action
- reader.cmd_v5_validate_plan — required_unfilled error + overflow:shrink warning
- reader.cmd_v5_compose — end-to-end build + shape placement
- Re-ingest preservation of user-set status / categories / overlap_decision
"""
from __future__ import annotations

import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "authoring"))
sys.path.insert(0, str(REPO / "consumer"))

from pptx import Presentation  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
from pptx.util import Inches  # noqa: E402

import ingest_v5  # noqa: E402
import reader as reader_mod  # noqa: E402


# ---------------------------------------------------------------------------
# Synthetic deck helpers
# ---------------------------------------------------------------------------


def make_title_body_deck(n_slides: int = 1):
    """Build an in-memory deck with N title+body slides."""
    prs = Presentation()
    for i in range(n_slides):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"Title {i+1}"
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 1:
                ph.text = f"Body line A {i+1}\nBody line B {i+1}"
    return prs


def stub_theme_v4():
    """Minimal v4 theme dict matching what extract_deck_theme produces."""
    return {
        "deck": "synthetic",
        "palette": {"dk1": "#000000", "lt1": "#FFFFFF", "accent1": "#FF0000"},
        "aliases": {"primary": "accent1", "accent": "accent1",
                    "text": "dk1", "background": "lt1"},
        "fonts": {"major": "Calibri", "minor": "Calibri"},
    }


def add_picture_to_slide(slide, png_bytes: bytes, left, top, width, height):
    """Add a Picture shape from raw PNG bytes."""
    return slide.shapes.add_picture(io.BytesIO(png_bytes), left, top, width, height)


# Tiny valid PNG (1x1 black pixel) — minimal binary for test pictures.
_TINY_PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452"
    "00000001000000010802000000907753"
    "de0000000c4944415478da6300010000"
    "0500010d0a2db40000000049454e44ae"
    "426082"
)


# ---------------------------------------------------------------------------
# digest_skeleton — slot inventory
# ---------------------------------------------------------------------------


class TestDigestSkeleton(unittest.TestCase):
    def test_writes_skeleton_yaml_with_slots(self):
        prs = make_title_body_deck(1)
        with tempfile.TemporaryDirectory() as tmp:
            skel_root = Path(tmp) / "skeletons"
            out = ingest_v5.digest_skeleton(
                prs.slides[0], prs.slide_width, prs.slide_height,
                "test", 1, stub_theme_v4(), skel_root,
            )
        self.assertEqual(out["id"], "test_01")
        self.assertEqual(out["status"], "pending")
        self.assertTrue(any(s["kind"] == "heading" for s in out["slots"]))
        self.assertTrue(any(s["kind"] == "bullets" for s in out["slots"]))

    def test_every_slot_has_shape_id(self):
        prs = make_title_body_deck(1)
        with tempfile.TemporaryDirectory() as tmp:
            out = ingest_v5.digest_skeleton(
                prs.slides[0], prs.slide_width, prs.slide_height,
                "test", 1, stub_theme_v4(), Path(tmp),
            )
        for slot in out["slots"]:
            self.assertIn("shape_id", slot,
                          f"slot {slot.get('id')} missing shape_id (re-ingest preservation breaks without it)")

    def test_unmapped_shapes_captures_skipped(self):
        prs = make_title_body_deck(1)
        # Add a freeform (auto-shape) that the heuristic drops.
        slide = prs.slides[0]
        slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(0.1), Inches(0.1), Inches(0.3), Inches(0.3),
        )
        with tempfile.TemporaryDirectory() as tmp:
            out = ingest_v5.digest_skeleton(
                slide, prs.slide_width, prs.slide_height,
                "test", 1, stub_theme_v4(), Path(tmp),
            )
        self.assertGreater(len(out["unmapped_shapes"]), 0)
        kinds = {u["kind_hint"] for u in out["unmapped_shapes"]}
        self.assertIn("auto_shape", kinds)


# ---------------------------------------------------------------------------
# compute_repeated_picture_info — brand-mark detection
# ---------------------------------------------------------------------------


class TestRepeatedPictureDetection(unittest.TestCase):
    def test_picture_on_every_slide_is_flagged(self):
        prs = make_title_body_deck(3)
        # Add the same picture to all 3 slides at the same small size.
        for slide in prs.slides:
            add_picture_to_slide(
                slide, _TINY_PNG,
                Inches(0.1), Inches(0.1), Inches(0.5), Inches(0.5),
            )
        info = ingest_v5.compute_repeated_picture_info(prs)
        self.assertEqual(len(info), 1)
        median = list(info.values())[0]
        self.assertGreater(median, 0)

    def test_unique_pictures_not_flagged(self):
        prs = make_title_body_deck(3)
        # Different bytes per slide → different SHAs.
        for i, slide in enumerate(prs.slides):
            png = _TINY_PNG + bytes([i])  # vary by 1 byte
            add_picture_to_slide(
                slide, png,
                Inches(0.1), Inches(0.1), Inches(0.5), Inches(0.5),
            )
        info = ingest_v5.compute_repeated_picture_info(prs)
        self.assertEqual(info, {})


# ---------------------------------------------------------------------------
# Auto-classifier
# ---------------------------------------------------------------------------


class TestCategoryProposal(unittest.TestCase):
    def test_table_slot_triggers_data_category(self):
        slots = [
            {"kind": "heading", "geometry": {"h": 0.1}},
            {"kind": "table", "geometry": {"h": 0.5}},
        ]
        cats = ingest_v5._propose_categories(slots)
        self.assertIn("data", cats)

    def test_closing_pattern_in_text(self):
        slots = [
            {"kind": "heading", "geometry": {"h": 0.2},
             "source_excerpt": "Thank you for your attention"},
        ]
        cats = ingest_v5._propose_categories(slots)
        self.assertIn("closing", cats)

    def test_fallback_to_content(self):
        slots = [{"kind": "paragraph", "geometry": {"h": 0.5},
                  "source_excerpt": "Some paragraph text"}]
        cats = ingest_v5._propose_categories(slots)
        self.assertEqual(cats, ["content"])


# ---------------------------------------------------------------------------
# Re-ingest preservation
# ---------------------------------------------------------------------------


class TestReingestPreservation(unittest.TestCase):
    def test_user_set_status_survives(self):
        prs = make_title_body_deck(1)
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            ingest_v5.digest_skeleton(
                prs.slides[0], prs.slide_width, prs.slide_height,
                "deckA", 1, stub_theme_v4(), root,
            )
            # User marks done + sets overlap_decision via the UI
            sk_path = root / "deckA_01" / "skeleton.yaml"
            import yaml
            data = yaml.safe_load(sk_path.read_text())
            data["status"] = "done"
            data["overlap_decision"] = "image_slot"
            data["categories"] = ["data"]
            sk_path.write_text(yaml.safe_dump(data, sort_keys=False))
            # Re-ingest
            out = ingest_v5.digest_skeleton(
                prs.slides[0], prs.slide_width, prs.slide_height,
                "deckA", 1, stub_theme_v4(), root,
            )
        self.assertEqual(out["status"], "done")
        self.assertEqual(out["overlap_decision"], "image_slot")
        self.assertEqual(out["categories"], ["data"])


# ---------------------------------------------------------------------------
# Constraint helpers
# ---------------------------------------------------------------------------


class TestConstraintHelpers(unittest.TestCase):
    def test_text_fit_under(self):
        fits, _, hr = reader_mod._v5_check_text_fit("hello", {"max_chars": 10})
        self.assertTrue(fits)
        self.assertEqual(hr, 5)

    def test_text_fit_over(self):
        fits, reason, hr = reader_mod._v5_check_text_fit("hello world", {"max_chars": 5})
        self.assertFalse(fits)
        self.assertLess(hr, 0)
        self.assertIn("max_chars", reason)

    def test_bullets_too_many_items(self):
        fits, reason, _ = reader_mod._v5_check_bullets_fit(["a", "b", "c"], {"max_items": 2})
        self.assertFalse(fits)
        self.assertIn("max_items", reason)


# ---------------------------------------------------------------------------
# Reader CLI commands — exercised via cmd_* functions directly to stdout
# ---------------------------------------------------------------------------


class _StubArgs:
    """Minimal stand-in for argparse.Namespace."""
    def __init__(self, **kw):
        self.__dict__.update(kw)


def _capture_stdout(fn, *args):
    """Run a cmd_v5_* function; capture and parse its stdout JSON."""
    buf = io.StringIO()
    real_stdout = sys.stdout
    sys.stdout = buf
    try:
        fn(*args)
    finally:
        sys.stdout = real_stdout
    return json.loads(buf.getvalue())


class TestMatchSkeletons(unittest.TestCase):
    def setUp(self):
        # Build a workspace and patch _v5_bundle_root.
        self.tmpdir = tempfile.TemporaryDirectory()
        root = Path(self.tmpdir.name)
        (root / "themes").mkdir()
        (root / "skeletons").mkdir()
        # Two skeletons: tight title (max 30) vs loose (max 200)
        for sid, max_chars in [("tight_01", 30), ("loose_01", 200)]:
            d = root / "skeletons" / sid
            d.mkdir()
            (d / "skeleton.yaml").write_text(
                f"id: {sid}\n"
                f"source_deck: synth\n"
                f"source_slide_index: 1\n"
                f"status: pending\n"
                f"categories: [content]\n"
                f"slots:\n"
                f"  - id: title\n"
                f"    kind: heading\n"
                f"    geometry: {{x: 0.05, y: 0.05, w: 0.9, h: 0.1}}\n"
                f"    constraints: {{max_chars: {max_chars}, required: true}}\n",
            )
        self._patcher = mock.patch.object(reader_mod, "_v5_bundle_root",
                                          return_value=root)
        self._patcher.start()

    def tearDown(self):
        self._patcher.stop()
        self.tmpdir.cleanup()

    def test_tighter_skeleton_ranks_higher(self):
        out = _capture_stdout(
            reader_mod.cmd_v5_match_skeletons,
            _StubArgs(content='{"title": "Q4 results beat consensus"}',
                      category=None, has_slot=None),
        )
        self.assertEqual(len(out["matches"]), 2)
        # The 30-char-max skeleton fits a 27-char title more tightly
        # than the 200-char-max one → ranks first.
        self.assertEqual(out["matches"][0]["skeleton_id"], "tight_01")

    def test_zero_match_returns_suggested_action(self):
        long_title = "X" * 250  # exceeds both 30 and 200 char limits
        out = _capture_stdout(
            reader_mod.cmd_v5_match_skeletons,
            _StubArgs(content=json.dumps({"title": long_title}),
                      category=None, has_slot=None),
        )
        self.assertEqual(out["matches"], [])
        self.assertGreater(len(out["issues"]), 0)
        issue = out["issues"][0]
        self.assertIn("suggested_action", issue)
        # Tightest constraint across both candidates is 30
        self.assertEqual(issue["tightest_constraint"], 30)


class TestValidatePlan(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        root = Path(self.tmpdir.name)
        (root / "themes").mkdir()
        (root / "skeletons" / "sk_01").mkdir(parents=True)
        (root / "skeletons" / "sk_01" / "skeleton.yaml").write_text(
            "id: sk_01\n"
            "source_deck: synth\n"
            "source_slide_index: 1\n"
            "status: pending\n"
            "categories: [content]\n"
            "slots:\n"
            "  - id: title\n"
            "    kind: heading\n"
            "    geometry: {x: 0, y: 0, w: 1, h: 0.2}\n"
            "    constraints: {max_chars: 20, required: true}\n"
            "  - id: body\n"
            "    kind: paragraph\n"
            "    geometry: {x: 0, y: 0.3, w: 1, h: 0.5}\n"
            "    constraints: {max_chars: 100, required: false}\n",
        )
        self._patcher = mock.patch.object(reader_mod, "_v5_bundle_root",
                                          return_value=root)
        self._patcher.start()
        self.root = root

    def tearDown(self):
        self._patcher.stop()
        self.tmpdir.cleanup()

    def test_required_unfilled_is_error(self):
        plan = self.root / "plan.json"
        plan.write_text(json.dumps([{"skeleton_id": "sk_01", "slots": {"body": "ok"}}]))
        out = _capture_stdout(
            reader_mod.cmd_v5_validate_plan,
            _StubArgs(plan=str(plan)),
        )
        self.assertFalse(out["ok"])
        self.assertTrue(any(e["violation"] == "required_unfilled" for e in out["errors"]))

    def test_overflow_shrink_is_warning_not_error(self):
        plan = self.root / "plan.json"
        plan.write_text(json.dumps([{
            "skeleton_id": "sk_01",
            "slots": {
                "title": {"value": "Way too long for the 20-char title slot",
                          "overflow": "shrink"},
            },
        }]))
        out = _capture_stdout(
            reader_mod.cmd_v5_validate_plan,
            _StubArgs(plan=str(plan)),
        )
        self.assertGreater(len(out["warnings"]), 0)
        warning_slots = [w["slot_id"] for w in out["warnings"]]
        self.assertIn("title", warning_slots)
        # Title isn't a hard error because overflow:shrink covers it
        error_slots = [e.get("slot_id") for e in out["errors"]]
        self.assertNotIn("title", error_slots)


# ---------------------------------------------------------------------------
# Compose round-trip
# ---------------------------------------------------------------------------


class TestComposeRoundTrip(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tmpdir.name)
        (self.root / "skeletons" / "sk_01").mkdir(parents=True)
        # Write a real master.pptx
        themes = self.root / "themes" / "synth"
        themes.mkdir(parents=True)
        synth_prs = make_title_body_deck(1)
        synth_prs.save(str(themes / "master.pptx"))
        (themes / "theme.yaml").write_text(
            "id: synth\n"
            "palette: {primary: '#FF0000', accent: '#00FF00'}\n"
            "fonts: {major: Calibri, minor: Calibri}\n"
            "master_pptx: master.pptx\n",
        )
        (self.root / "skeletons" / "sk_01" / "skeleton.yaml").write_text(
            "id: sk_01\n"
            "source_deck: synth\n"
            "source_slide_index: 1\n"
            "status: pending\n"
            "categories: [content]\n"
            "slots:\n"
            "  - id: title\n"
            "    kind: heading\n"
            "    geometry: {x: 0.05, y: 0.05, w: 0.9, h: 0.15}\n"
            "    style: {font_role: major, size_pt: 32}\n"
            "    constraints: {max_chars: 50, required: true}\n"
            "  - id: body\n"
            "    kind: bullets\n"
            "    geometry: {x: 0.05, y: 0.3, w: 0.9, h: 0.6}\n"
            "    constraints: {max_items: 5, max_chars_per_item: 80}\n",
        )
        self._patcher = mock.patch.object(reader_mod, "_v5_bundle_root",
                                          return_value=self.root)
        self._patcher.start()

    def tearDown(self):
        self._patcher.stop()
        self.tmpdir.cleanup()

    def test_compose_produces_real_pptx(self):
        plan = self.root / "plan.json"
        plan.write_text(json.dumps([{
            "skeleton_id": "sk_01",
            "slots": {"title": "Hello world", "body": ["item a", "item b"]},
        }]))
        out_pptx = self.root / "out.pptx"
        _capture_stdout(
            reader_mod.cmd_v5_compose,
            _StubArgs(plan=str(plan), out=str(out_pptx), theme="synth"),
        )
        self.assertTrue(out_pptx.exists())
        prs = Presentation(str(out_pptx))
        self.assertEqual(len(prs.slides), 1)
        # Two text boxes added (title + body bullets)
        text_boxes = [s for s in prs.slides[0].shapes
                      if s.has_text_frame and s.text_frame.text.strip()]
        all_text = "\n".join(s.text_frame.text for s in text_boxes)
        self.assertIn("Hello world", all_text)
        self.assertIn("item a", all_text)


if __name__ == "__main__":
    unittest.main()
