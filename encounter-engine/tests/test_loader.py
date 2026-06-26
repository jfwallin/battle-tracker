"""Tests for excel_loader against the real combat tracker.xlsx."""
import os
import sys
import pytest

# Make sure the package root is on the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from engine.excel_loader import load_encounter, parse_damage, parse_to_hit, parse_save

XLSX = os.path.join(
    os.path.dirname(__file__),
    "..", "..", "combat tracker.xlsx"
)


def test_xlsx_exists():
    assert os.path.exists(XLSX), f"Excel file not found at {XLSX}"


def test_load_encounter_returns_combatants():
    enc = load_encounter(XLSX)
    assert len(enc.combatants) > 0, "Expected at least one combatant"


def test_combatant_names():
    enc = load_encounter(XLSX)
    names = [c.name for c in enc.combatants.values()]
    # At minimum we expect entries for the two creature sheets
    assert any("Korrum" in n or "Veyrath" in n for n in names), (
        f"Expected Korrum or Veyrath in combatant names, got: {names}"
    )


def test_korrum_stats():
    enc = load_encounter(XLSX)
    korrum = next(
        (c for c in enc.combatants.values() if "Korrum" in c.name), None
    )
    assert korrum is not None, "Korrum not found"
    assert korrum.ac is not None, "Korrum AC should be set"
    assert korrum.max_hp is not None and korrum.max_hp > 0, "Korrum max HP should be > 0"
    assert len(korrum.attacks) > 0, "Korrum should have at least one attack"


def test_veyrath_stats():
    enc = load_encounter(XLSX)
    veyrath = next(
        (c for c in enc.combatants.values() if "Veyrath" in c.name), None
    )
    assert veyrath is not None, "Veyrath not found"
    assert veyrath.ac is not None, "Veyrath AC should be set"
    assert veyrath.max_hp is not None and veyrath.max_hp > 0, "Veyrath max HP should be > 0"
    assert len(veyrath.attacks) > 0, "Veyrath should have at least one attack"


def test_parse_to_hit():
    assert parse_to_hit("+10") == 10
    assert parse_to_hit("10") == 10
    assert parse_to_hit("-2") == -2
    assert parse_to_hit("—") is None
    assert parse_to_hit("") is None


def test_parse_damage_simple():
    result = parse_damage("2d8+6 fire")
    assert len(result) == 1
    d = result[0]
    assert d.n == 2
    assert d.die == 8
    assert d.bonus == 6
    assert d.damage_type == "fire"


def test_parse_damage_multi():
    result = parse_damage("1d8+3 slashing plus 2d6 fire")
    assert len(result) == 2
    assert result[0].damage_type == "slashing"
    assert result[1].damage_type == "fire"


def test_parse_damage_no_type():
    result = parse_damage("3d10+7")
    assert len(result) == 1
    assert result[0].n == 3
    assert result[0].die == 10
    assert result[0].bonus == 7


def test_parse_save():
    dc, ability = parse_save("DC 17 STR")
    assert dc == 17
    assert ability == "STR"
    dc2, ability2 = parse_save("—")
    assert dc2 is None
    assert ability2 == ""


def test_order_matches_combatants():
    enc = load_encounter(XLSX)
    for cid in enc.order:
        assert cid in enc.combatants, f"Order id {cid} not in combatants"
