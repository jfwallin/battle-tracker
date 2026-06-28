"""Tests for the Battle Library (engine/battle_library.py) — cross-workbook resolution."""
import os
import shutil
import sys

import pytest
import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from engine.battle_library import list_battles, load_battle
from tests.make_sample_library import make_sample_library, _DEF_HEADER, _INDEX_HEADER

REAL_TRACKER = os.path.join(os.path.dirname(__file__), "..", "..", "combat tracker.xlsx")


@pytest.fixture
def tracker(tmp_path):
    """A copy of the real Combat Tracker (copy works even if it's open in Excel)."""
    if not os.path.exists(REAL_TRACKER):
        pytest.skip("combat tracker.xlsx not found")
    dst = str(tmp_path / "Combat Tracker.xlsx")
    shutil.copy2(REAL_TRACKER, dst)
    return dst


@pytest.fixture
def library(tmp_path):
    return make_sample_library(str(tmp_path / "Battle_Library.xlsx"))


def test_list_battles_counts(library, tracker):
    battles = list_battles(library, tracker)
    assert len(battles) == 1
    b = battles[0]
    assert b["id"] == "BTL-001"
    assert b["name"] == "Test Skirmish"
    assert b["status"] == "Ready"
    assert b["tags"] == ["test"]
    # Korrum (NPC x1) + Veyrath (NPC x2, group)
    assert b["counts"]["npcs"] == 2
    assert b["counts"]["groups"] == 1
    assert not b["warnings"]  # valid sources, valid modes


def test_load_battle_resolves_cross_workbook(library, tracker):
    enc = load_battle(library, tracker, "BTL-001")
    assert len(enc.order) == 2
    by_name = {c.name: c for c in enc.combatants.values()}

    korrum = by_name["Korrum"]
    assert korrum.ac is not None and korrum.max_hp and korrum.max_hp > 0
    assert len(korrum.attacks) > 0  # stats pulled from the OTHER workbook

    veyrath = by_name["Veyrath Pack"]  # group_name overrides display name
    assert veyrath.is_group and len(veyrath.members) == 2
    assert veyrath.initiative_mode == "Mob"
    # Starting condition from col I applied.
    assert "Frightened" in veyrath.conditions

    # source_path points at the tracker so export can still write back.
    assert enc.source_path == tracker


def test_broken_source_tab_warns(tmp_path, tracker):
    # Hand-build a library whose definition references a non-existent creature tab.
    wb = openpyxl.Workbook()
    idx = wb.active
    idx.title = "Battle Index"
    idx.append(_INDEX_HEADER)
    idx.append(["BTL-X", "Broken", "BTL-X Def", "Draft", "", "", "", ""])
    d = wb.create_sheet("BTL-X Def")
    d.append(_DEF_HEADER)
    d.append(["NPC", "Ghost", "NoSuchTab", 1, "", "Individual", "", "", "", "", "", ""])
    lib = str(tmp_path / "Battle_Library.xlsx")
    wb.save(lib)

    battles = list_battles(lib, tracker)
    assert any("NoSuchTab" in w for w in battles[0]["warnings"])

    # load_battle still builds the combatant, just without stats (graceful).
    enc = load_battle(lib, tracker, "BTL-X")
    ghost = next(iter(enc.combatants.values()))
    assert ghost.name == "Ghost"
    assert ghost.ac is None and ghost.max_hp is None


def test_unknown_battle_raises(library, tracker):
    with pytest.raises(ValueError):
        load_battle(library, tracker, "NOPE")
