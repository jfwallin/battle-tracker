"""Tests for reusable variants (engine/game_data.py) + battle-load integration."""
import os
import shutil
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from engine.game_data import (
    _apply_expr, apply_variant, delete_variant, get_variant, list_variants, save_variant,
)
from engine.battle_library import load_battle, save_battle
from engine.models import Combatant, Member, Attack, DamageDie, make_id

REAL_TRACKER = os.path.join(os.path.dirname(__file__), "..", "..", "combat tracker.xlsx")


def test_apply_expr():
    assert _apply_expr(15, "20") == 20      # absolute
    assert _apply_expr(15, "+3") == 18      # delta
    assert _apply_expr(15, "-2") == 13
    assert _apply_expr(10, "x1.5") == 15    # multiply
    assert _apply_expr(10, "*0.5") == 5
    assert _apply_expr(15, "") == 15        # blank = unchanged
    assert _apply_expr(None, "+3") is None  # delta on missing = skip
    assert _apply_expr(None, "18") == 18    # absolute can set from None


def _make_combatant():
    return Combatant(
        id=make_id(), name="Hobgoblin", combatant_type="NPC", source_tab="Hobgoblin",
        initiative=None, initiative_mod=2, ac=18, max_hp=11,
        attacks=[Attack(name="Longbow", attack_type="Ranged", to_hit=3, reach="150 ft.",
                        damage_dice=[DamageDie(n=1, die=8, bonus=1, damage_type="piercing")])],
        resistances="", immunities="",
    )


def test_apply_variant_scalars_and_attacks():
    c = _make_combatant()
    apply_variant(c, {"name": "Elite", "ac": "+2", "max_hp": "x2", "to_hit": "+2",
                      "resistances": "fire", "immunities": "poison", "notes": "tougher"})
    assert c.name == "Hobgoblin (Elite)"
    assert c.ac == 20 and c.max_hp == 22
    assert c.attacks[0].to_hit == 5
    assert "fire" in c.resistances and "poison" in c.immunities
    assert "tougher" in c.notes


def test_apply_variant_group_syncs_members():
    g = Combatant(id=make_id(), name="Archers", combatant_type="NPC", source_tab="Hobgoblin",
                  initiative=None, initiative_mod=2, ac=18, max_hp=11, is_group=True,
                  members=[Member(id=make_id(), name=f"A{i}", max_hp=11) for i in range(3)])
    apply_variant(g, {"name": "Veteran", "max_hp": "+10"})
    assert g.max_hp == 21
    assert all(m.max_hp == 21 for m in g.members)


def test_save_list_get_delete_variant(tmp_path):
    gd = str(tmp_path / "Game_Data.xlsx")
    save_variant(gd, {"id": "elite", "name": "Elite", "ac": "+2", "max_hp": "x1.5"})
    save_variant(gd, {"id": "wounded", "name": "Wounded", "max_hp": "x0.5"})
    ids = {v["id"] for v in list_variants(gd)}
    assert ids == {"elite", "wounded"}
    assert get_variant(gd, "elite")["ac"] == "+2"

    # Re-save replaces, not appends.
    save_variant(gd, {"id": "elite", "name": "Elite+", "ac": "+3"})
    assert len(list_variants(gd)) == 2 and get_variant(gd, "elite")["name"] == "Elite+"

    delete_variant(gd, "wounded")
    assert {v["id"] for v in list_variants(gd)} == {"elite"}


def test_load_battle_applies_variant(tmp_path):
    if not os.path.exists(REAL_TRACKER):
        pytest.skip("combat tracker.xlsx not found")
    tracker = str(tmp_path / "Combat Tracker.xlsx")
    shutil.copy2(REAL_TRACKER, tracker)
    gd = str(tmp_path / "Game_Data.xlsx")
    save_variant(gd, {"id": "elite", "name": "Elite", "ac": "+2", "max_hp": "+50"})

    lib = str(tmp_path / "Battle_Library.xlsx")
    save_battle(lib, {"name": "VarTest", "rows": [
        {"type": "NPC", "name": "Boss", "npc_source": "Korrum", "quantity": 1, "variant_id": "elite"},
        {"type": "NPC", "name": "Plain", "npc_source": "Veyrath", "quantity": 1},
    ]})

    enc = load_battle(lib, tracker, "BTL-001", game_data_path=gd)
    by_name = {c.name: c for c in enc.combatants.values()}
    boss = by_name["Boss (Elite)"]   # variant suffix applied
    assert boss.ac == 20            # Korrum AC 18 + 2
    assert boss.max_hp == 335       # Korrum HP 285 + 50
    plain = by_name["Plain"]         # un-varianted creature unaffected
    assert plain.ac == 17 and plain.source_tab == "Veyrath"  # Veyrath's own AC, no suffix
