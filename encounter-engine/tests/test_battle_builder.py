"""Tests for Stage 2 builder backend: list_sources + save/edit/duplicate/delete round-trip."""
import os
import shutil
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from engine.battle_library import (
    delete_battle, duplicate_battle, get_battle_definition, list_battles,
    list_sources, load_battle, save_battle,
)

REAL_TRACKER = os.path.join(os.path.dirname(__file__), "..", "..", "combat tracker.xlsx")


@pytest.fixture
def tracker(tmp_path):
    if not os.path.exists(REAL_TRACKER):
        pytest.skip("combat tracker.xlsx not found")
    dst = str(tmp_path / "Combat Tracker.xlsx")
    shutil.copy2(REAL_TRACKER, dst)
    return dst


def test_list_sources(tracker):
    sources = list_sources(tracker)
    names = {s["name"] for s in sources}
    assert "Korrum" in names and "Veyrath" in names
    assert "Battle List" not in {s["tab"] for s in sources}
    korrum = next(s for s in sources if s["name"] == "Korrum")
    assert korrum["ac"] == 18 and korrum["max_hp"] == 285


def test_save_creates_library_then_loads(tmp_path, tracker):
    lib = str(tmp_path / "Battle_Library.xlsx")  # does not exist yet
    saved = save_battle(lib, {
        "name": "Goblin Skirmish",
        "status": "Ready",
        "tags": ["test", "goblins"],
        "notes": "A quick fight",
        "rows": [
            {"type": "NPC", "name": "Hob", "npc_source": "Hobgoblin", "quantity": 3,
             "group_name": "Reds", "initiative_mode": "Mob"},
            {"type": "Hazard", "name": "Rockfall", "quantity": 1, "initiative_mode": "Individual",
             "notes": "DC 13 Dex"},
        ],
    })
    assert os.path.exists(lib)
    assert saved["id"] == "BTL-001"

    battles = list_battles(lib, tracker)
    assert len(battles) == 1
    b = battles[0]
    assert b["counts"]["npcs"] == 1 and b["counts"]["groups"] == 1 and b["counts"]["hazards"] == 1
    assert not b["warnings"]

    # Loads into a live encounter with stats resolved from the tracker.
    enc = load_battle(lib, tracker, "BTL-001")
    hob = next(c for c in enc.combatants.values() if c.is_group)
    assert len(hob.members) == 3 and hob.ac == 18  # Hobgoblin AC from tracker


def test_edit_preserves_sheet_and_updates_index(tmp_path, tracker):
    lib = str(tmp_path / "Battle_Library.xlsx")
    saved = save_battle(lib, {"name": "Edit Me", "rows": [
        {"type": "NPC", "name": "K", "npc_source": "Korrum", "quantity": 1}]})
    defn = get_battle_definition(lib, saved["id"])
    loc = defn["definition_location"]

    # Edit: rename + add a row, keeping the same id/sheet.
    save_battle(lib, {
        "id": saved["id"], "definition_location": loc, "name": "Edited Name",
        "status": "Ready",
        "rows": defn["rows"] + [{"type": "NPC", "name": "V", "npc_source": "Veyrath", "quantity": 2}],
    })
    after = get_battle_definition(lib, saved["id"])
    assert after["name"] == "Edited Name"
    assert after["definition_location"] == loc  # same sheet reused
    assert len(after["rows"]) == 2
    # Still exactly one battle in the index (updated, not appended).
    assert len(list_battles(lib, tracker)) == 1


def test_duplicate_then_delete(tmp_path, tracker):
    lib = str(tmp_path / "Battle_Library.xlsx")
    a = save_battle(lib, {"name": "Original", "rows": [
        {"type": "NPC", "name": "K", "npc_source": "Korrum", "quantity": 1}]})
    dup = duplicate_battle(lib, a["id"])
    assert dup["id"] != a["id"]
    names = {b["name"] for b in list_battles(lib, tracker)}
    assert "Original" in names and "Original (copy)" in names
    assert dup["definition_location"] != a["definition_location"]  # separate sheet

    delete_battle(lib, dup["id"])
    remaining = list_battles(lib, tracker)
    assert len(remaining) == 1 and remaining[0]["id"] == a["id"]


def test_save_requires_name(tmp_path):
    lib = str(tmp_path / "Battle_Library.xlsx")
    with pytest.raises(ValueError):
        save_battle(lib, {"name": "  ", "rows": []})
