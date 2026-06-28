"""Tests for party profiles (engine/game_data.py)."""
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from engine.game_data import (
    apply_party_profile, build_party_combatants, delete_party_profile,
    get_party_profile, list_party_profiles, save_party_profile,
)
from engine.models import Encounter, Combatant


def _profile(pid, members):
    return {"id": pid, "members": members}


def test_save_creates_file_and_lists(tmp_path):
    gd = str(tmp_path / "Game_Data.xlsx")
    save_party_profile(gd, _profile("PARTY-MAIN", [
        {"name": "Rex", "type": "PC", "ac": 17, "max_hp": 45, "init_mod": 2, "notes": "tank"},
        {"name": "Giggles", "type": "PC", "ac": 13, "max_hp": 28, "init_mod": 4},
    ]))
    assert os.path.exists(gd)
    profiles = list_party_profiles(gd)
    assert len(profiles) == 1
    p = profiles[0]
    assert p["id"] == "PARTY-MAIN" and p["count"] == 2
    rex = p["members"][0]
    assert rex["name"] == "Rex" and rex["ac"] == 17 and rex["max_hp"] == 45 and rex["init_mod"] == 2


def test_build_party_combatants(tmp_path):
    gd = str(tmp_path / "Game_Data.xlsx")
    save_party_profile(gd, _profile("P", [
        {"name": "Ally", "type": "Ally", "ac": 15, "max_hp": 30, "init_mod": 1}]))
    combs = build_party_combatants(get_party_profile(gd, "P"))
    assert len(combs) == 1
    c = combs[0]
    assert isinstance(c, Combatant)
    assert c.name == "Ally" and c.combatant_type == "Ally"
    assert c.ac == 15 and c.max_hp == 30 and c.initiative_mod == 1 and c.source_tab is None


def test_apply_party_profile_prepends(tmp_path):
    gd = str(tmp_path / "Game_Data.xlsx")
    save_party_profile(gd, _profile("PARTY", [
        {"name": "Rex", "type": "PC"}, {"name": "Prag", "type": "PC"}]))
    enc = Encounter(source_path="x")
    enemy = Combatant(id="e1", name="Goblin", combatant_type="NPC", source_tab=None,
                      initiative=None, initiative_mod=0, ac=15, max_hp=7)
    enc.combatants["e1"] = enemy
    enc.order = ["e1"]

    added = apply_party_profile(enc, gd, "PARTY")
    assert added == 2
    assert len(enc.order) == 3
    # Party leads the roster; enemy stays last.
    assert [enc.combatants[i].name for i in enc.order[:2]] == ["Rex", "Prag"]
    assert enc.combatants[enc.order[-1]].name == "Goblin"


def test_edit_and_delete_preserve_other_profiles(tmp_path):
    gd = str(tmp_path / "Game_Data.xlsx")
    save_party_profile(gd, _profile("A", [{"name": "A1"}]))
    save_party_profile(gd, _profile("B", [{"name": "B1"}, {"name": "B2"}]))
    assert {p["id"] for p in list_party_profiles(gd)} == {"A", "B"}

    # Re-save A with different members — replaces A's rows, keeps B intact.
    save_party_profile(gd, _profile("A", [{"name": "A1"}, {"name": "A2"}, {"name": "A3"}]))
    profiles = {p["id"]: p for p in list_party_profiles(gd)}
    assert profiles["A"]["count"] == 3 and profiles["B"]["count"] == 2

    delete_party_profile(gd, "A")
    remaining = list_party_profiles(gd)
    assert len(remaining) == 1 and remaining[0]["id"] == "B"


def test_save_requires_id(tmp_path):
    gd = str(tmp_path / "Game_Data.xlsx")
    with pytest.raises(ValueError):
        save_party_profile(gd, _profile("  ", [{"name": "X"}]))
