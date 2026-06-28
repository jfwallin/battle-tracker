"""Workspace directory discovery and the remembered-workspace config.

A *workspace* is a directory holding the standard encounter files. Stage 1 cares about
three filenames; only Combat Tracker is required. Filename matching is case-insensitive
because the real library file is ``combat tracker.xlsx`` (lowercase) on disk.

This module is pure (no Flask) so it can be unit-tested directly.
"""
from __future__ import annotations
import json
import os
from datetime import datetime, timezone
from typing import Optional

# data/workspace.json lives next to encounter_state.json (encounter-engine/data/).
WORKSPACE_CONFIG_PATH = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "workspace.json"
)

# key -> expected standard filename
STANDARD_FILES = {
    "combat_tracker": "Combat Tracker.xlsx",
    "battle_library": "Battle_Library.xlsx",
    "game_data": "Game_Data.xlsx",   # optional
}

# Per-file message templates: (found_message, missing_message)
_MESSAGES = {
    "combat_tracker": (
        "Combat Tracker found: {filename}.",
        "Combat Tracker.xlsx not found in this directory; creatures cannot be loaded.",
    ),
    "battle_library": (
        "Battle Library found: {filename}.",
        "Battle Library not found; prepared battles unavailable. "
        "You can still load a single workbook manually.",
    ),
    "game_data": (
        "Game Data found: {filename}.",
        "Game Data not found; continuing without party profiles, recurring NPCs, "
        "or reusable variants.",
    ),
}


def _find_file(workspace_dir: str, expected: str) -> Optional[str]:
    """Return the actual filename in workspace_dir matching `expected` (case-insensitive)."""
    target = expected.lower()
    try:
        for entry in os.listdir(workspace_dir):
            if entry.lower() == target:
                return entry
    except OSError:
        return None
    return None


def discover_workspace(workspace_dir: str) -> dict:
    """Inspect a directory for the standard files.

    Returns ``{dir, exists, files:{key:{found,path,filename,expected}}, messages:[...], ok}``.
    ``ok`` is True only when the required Combat Tracker is present. Optional files missing
    is reported but does not fail discovery.
    """
    result: dict = {"dir": workspace_dir, "exists": os.path.isdir(workspace_dir),
                    "files": {}, "messages": [], "ok": False}

    if not result["exists"]:
        result["messages"].append(f"Directory not found: {workspace_dir}")
        for key, expected in STANDARD_FILES.items():
            result["files"][key] = {"found": False, "path": None,
                                    "filename": None, "expected": expected}
        return result

    for key, expected in STANDARD_FILES.items():
        actual = _find_file(workspace_dir, expected)
        found = actual is not None
        result["files"][key] = {
            "found": found,
            "path": os.path.join(workspace_dir, actual) if found else None,
            "filename": actual,
            "expected": expected,
        }
        found_msg, missing_msg = _MESSAGES[key]
        result["messages"].append(found_msg.format(filename=actual) if found else missing_msg)

    result["ok"] = result["files"]["combat_tracker"]["found"]
    return result


def load_workspace_config() -> Optional[dict]:
    """Return the saved workspace config dict, or None if not set / unreadable."""
    if os.path.exists(WORKSPACE_CONFIG_PATH):
        try:
            with open(WORKSPACE_CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except (OSError, json.JSONDecodeError):
            return None
    return None


def save_workspace_config(workspace_dir: str) -> dict:
    """Persist the chosen workspace directory and return the stored config."""
    os.makedirs(os.path.dirname(WORKSPACE_CONFIG_PATH), exist_ok=True)
    config = {
        "workspace_dir": workspace_dir,
        "set_at": datetime.now(timezone.utc).isoformat(),
    }
    with open(WORKSPACE_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)
    return config
