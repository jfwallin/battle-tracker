"""Tests for workspace directory discovery (engine/workspace.py)."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from engine.workspace import discover_workspace, STANDARD_FILES


def _touch(path):
    with open(path, "w", encoding="utf-8") as f:
        f.write("")


def test_discover_all_present_case_insensitive(tmp_path):
    # Real tracker file is lowercase 'combat tracker.xlsx' — discovery must match it.
    _touch(tmp_path / "combat tracker.xlsx")
    _touch(tmp_path / "Battle_Library.xlsx")
    _touch(tmp_path / "Game_Data.xlsx")

    d = discover_workspace(str(tmp_path))
    assert d["exists"] is True
    assert d["ok"] is True
    assert d["files"]["combat_tracker"]["found"] is True
    assert d["files"]["combat_tracker"]["filename"] == "combat tracker.xlsx"
    assert d["files"]["battle_library"]["found"] is True
    assert d["files"]["game_data"]["found"] is True


def test_discover_only_tracker(tmp_path):
    _touch(tmp_path / "Combat Tracker.xlsx")
    d = discover_workspace(str(tmp_path))
    assert d["ok"] is True  # tracker present is enough
    assert d["files"]["battle_library"]["found"] is False
    assert d["files"]["game_data"]["found"] is False
    # Helpful, specific messages — not generic failures.
    joined = " ".join(d["messages"]).lower()
    assert "battle library not found" in joined
    assert "game data not found" in joined


def test_discover_missing_tracker_not_ok(tmp_path):
    _touch(tmp_path / "Battle_Library.xlsx")
    d = discover_workspace(str(tmp_path))
    assert d["ok"] is False
    assert d["files"]["combat_tracker"]["found"] is False


def test_discover_nonexistent_dir():
    d = discover_workspace(os.path.join("Z:\\", "definitely-not-here-12345"))
    assert d["exists"] is False
    assert d["ok"] is False
    assert any("not found" in m.lower() for m in d["messages"])


def test_standard_files_keys():
    assert set(STANDARD_FILES) == {"combat_tracker", "battle_library", "game_data"}
