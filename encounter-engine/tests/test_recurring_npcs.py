"""Tests for recurring NPCs (engine/game_data.py)."""
import os, sys
import pytest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from engine.game_data import (
    list_recurring_npcs, get_recurring_npc, save_recurring_npc, delete_recurring_npc,
)


def test_crud(tmp_path):
    gd = str(tmp_path / "Game_Data.xlsx")
    save_recurring_npc(gd, {"id": "voss", "name": "Captain Voss",
                            "base_source": "Ganador Veteran", "variant_id": "elite",
                            "notes": "holds a key", "status": "Active"})
    save_recurring_npc(gd, {"id": "mage", "name": "Court Mage", "base_source": "War Mage"})
    ids = {n["id"] for n in list_recurring_npcs(gd)}
    assert ids == {"voss", "mage"}
    voss = get_recurring_npc(gd, "voss")
    assert voss["name"] == "Captain Voss" and voss["base_source"] == "Ganador Veteran"
    assert voss["variant_id"] == "elite" and voss["notes"] == "holds a key"

    save_recurring_npc(gd, {"id": "voss", "name": "Captain Voss II", "base_source": "Ganador Veteran"})
    assert len(list_recurring_npcs(gd)) == 2
    assert get_recurring_npc(gd, "voss")["name"] == "Captain Voss II"

    delete_recurring_npc(gd, "mage")
    assert {n["id"] for n in list_recurring_npcs(gd)} == {"voss"}


def test_requires_id_and_base(tmp_path):
    gd = str(tmp_path / "Game_Data.xlsx")
    with pytest.raises(ValueError):
        save_recurring_npc(gd, {"id": "", "base_source": "X"})
    with pytest.raises(ValueError):
        save_recurring_npc(gd, {"id": "x", "base_source": ""})
