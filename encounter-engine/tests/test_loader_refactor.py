"""Regression: extracted roster helpers reproduce load_encounter on the real workbook."""
import os
import shutil
import sys

import pytest
import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from engine.excel_loader import (
    load_encounter, read_battle_list_rows, build_combatant_from_row,
    make_workbook_resolver,
)

REAL_TRACKER = os.path.join(os.path.dirname(__file__), "..", "..", "combat tracker.xlsx")


@pytest.fixture
def tracker(tmp_path):
    if not os.path.exists(REAL_TRACKER):
        pytest.skip("combat tracker.xlsx not found")
    dst = str(tmp_path / "combat tracker.xlsx")
    shutil.copy2(REAL_TRACKER, dst)
    return dst


def test_extracted_helpers_match_load_encounter(tracker):
    # Reference path: the (refactored) public loader.
    enc = load_encounter(tracker)
    ref = sorted(
        (c.name, c.combatant_type, c.ac, c.max_hp, c.is_group, len(c.members), len(c.attacks))
        for c in enc.combatants.values()
    )

    # Manual path: drive the extracted units directly.
    wb = openpyxl.load_workbook(tracker, data_only=True)
    rows = read_battle_list_rows(wb["Battle List"])
    resolve = make_workbook_resolver(wb)
    built = [build_combatant_from_row(r, resolve) for r in rows]
    built = [c for c in built if c is not None]
    manual = sorted(
        (c.name, c.combatant_type, c.ac, c.max_hp, c.is_group, len(c.members), len(c.attacks))
        for c in built
    )

    assert manual == ref
    assert len(manual) == len(enc.order)


def test_read_battle_list_rows_has_optional_columns(tracker):
    wb = openpyxl.load_workbook(tracker, data_only=True)
    rows = read_battle_list_rows(wb["Battle List"])
    assert rows, "expected at least one roster row"
    # Legacy sheet lacks I–L, but the keys must still be present (empty).
    for key in ("starting_conditions", "starting_status", "variant_id", "trigger"):
        assert key in rows[0]
