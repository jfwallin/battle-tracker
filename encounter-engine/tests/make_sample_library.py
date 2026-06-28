"""Generate a sample Battle_Library.xlsx for tests and manual smoke-testing.

The sample's definition references creature tabs that must exist in Combat Tracker.xlsx
(default: Korrum, Veyrath). Run directly to drop a library next to the tracker:

    python tests/make_sample_library.py "C:\\path\\to\\Battle_Library.xlsx"
"""
from __future__ import annotations
import sys

import openpyxl

# Battle Index header (cols A–H) and one prepared battle.
_INDEX_HEADER = [
    "Battle ID", "Battle Name", "Definition Location", "Status",
    "Tags", "Notes", "Default Party Profile", "Last Modified",
]
# Definition sheet header — identical to the live "Battle List" columns (A–H),
# plus the optional Stage-1 extras (I–L).
_DEF_HEADER = [
    "Type", "Name", "NPC Source", "Quantity", "Group Name",
    "Initiative Mode", "HP Override", "Notes",
    "Starting Conditions", "Starting Status", "Variant ID", "Trigger",
]


def make_sample_library(out_path: str) -> str:
    """Write a minimal Battle_Library.xlsx with one battle (BTL-001) and return out_path."""
    wb = openpyxl.Workbook()

    index = wb.active
    index.title = "Battle Index"
    index.append(_INDEX_HEADER)
    index.append([
        "BTL-001", "Test Skirmish", "BTL-001 Skirmish", "Ready",
        "test", "Smoke-test encounter", "", "",
    ])

    definition = wb.create_sheet("BTL-001 Skirmish")
    definition.append(_DEF_HEADER)
    # Korrum solo; Veyrath as a 2-member mob with a starting condition.
    definition.append(["NPC", "Korrum", "Korrum", 1, "", "Individual", "", "boss",
                       "", "", "", ""])
    definition.append(["NPC", "Veyrath", "Veyrath", 2, "Veyrath Pack", "Mob", "", "adds",
                       "Frightened", "", "", ""])

    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "Battle_Library.xlsx"
    print("Wrote", make_sample_library(target))
