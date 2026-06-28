"""Battle Library: reusable battle definitions that reference Combat Tracker creatures.

`Battle_Library.xlsx` has a ``Battle Index`` sheet (one row per prepared battle) plus one
definition sheet per battle, shaped exactly like the legacy ``Battle List`` so the same
row parser is reused. Creature/NPC sources in a definition resolve against the separate
``Combat Tracker.xlsx`` workbook.

Pure (no Flask) so it can be unit-tested directly.
"""
from __future__ import annotations
import os
import re
from datetime import date
from typing import Optional

import openpyxl
from openpyxl.styles import Font

from .models import Encounter
from .excel_loader import (
    ROSTER_COLUMNS,
    build_combatant_from_row,
    make_combat_tracker_resolver,
    read_battle_list_rows,
)

BATTLE_INDEX_SHEET = "Battle Index"
VALID_INIT_MODES = {"Individual", "Shared", "Grouped Attacks", "Mob"}

_INDEX_COLUMNS = [
    "battle_id", "name", "definition_location", "status", "tags",
    "notes", "default_party_profile", "last_modified",
]
_INDEX_HEADER = [
    "Battle ID", "Battle Name", "Definition Location", "Status",
    "Tags", "Notes", "Default Party Profile", "Last Modified",
]
# Definition-sheet header, in ROSTER_COLUMNS order (cols A–L).
_DEF_HEADER = [
    "Type", "Name", "NPC Source", "Quantity", "Group Name", "Initiative Mode",
    "HP Override", "Notes", "Starting Conditions", "Starting Status",
    "Variant ID", "Trigger",
]
_TYPE_TO_COUNT = {
    "pc": "pcs", "npc": "npcs", "ally": "allies",
    "hazard": "hazards", "event": "events",
}
_INVALID_SHEET_CHARS = r"[\\/?*\[\]:]"


def _open(path: str):
    return openpyxl.load_workbook(path, data_only=True)


def _read_index_rows(ws) -> list[dict]:
    """Read the Battle Index into row dicts (header found by 'Battle ID' in col A)."""
    header_row = None
    for row in ws.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        if cell.value and str(cell.value).strip().lower() in ("battle id", "battle_id"):
            header_row = cell.row
            break
    if header_row is None:
        header_row = 1

    rows: list[dict] = []
    for r in range(header_row + 1, ws.max_row + 1):
        def rv(col: int):
            v = ws.cell(row=r, column=col).value
            return str(v).strip() if v is not None else ""
        if not rv(1):
            continue  # blank Battle ID
        rows.append({key: rv(i + 1) for i, key in enumerate(_INDEX_COLUMNS)})
    return rows


def _count_and_validate(rows: list[dict], battle: dict, tracker_tabs: Optional[set]) -> None:
    """Fill battle['counts'] and append soft warnings for a definition's rows."""
    counts = battle["counts"]
    warnings = battle["warnings"]
    for row in rows:
        ctype = row.get("type", "")
        if not ctype:
            continue
        label = row.get("name", "") or row.get("group_name", "") or "?"

        qraw = row.get("quantity", "")
        try:
            qty = int(qraw) if qraw else 1
        except ValueError:
            qty = 1
            warnings.append(f"Row '{label}': quantity '{qraw}' invalid, treated as 1.")
        if qty <= 0:
            continue

        key = _TYPE_TO_COUNT.get(ctype.strip().lower())
        if key:
            counts[key] += 1
        if qty > 1:
            counts["groups"] += 1

        mode = row.get("initiative_mode", "")
        if mode and mode not in VALID_INIT_MODES:
            warnings.append(f"Row '{label}': initiative mode '{mode}' unknown; will use Individual.")

        src = row.get("npc_source", "")
        if src and tracker_tabs is not None and src not in tracker_tabs:
            warnings.append(f"Source tab '{src}' not found in Combat Tracker.")

        if row.get("variant_id"):
            warnings.append(f"Row '{label}': Variant ID not supported yet (Stage 3); ignored.")
        if row.get("trigger"):
            warnings.append(f"Row '{label}': Trigger not supported yet (Stage 5); ignored.")


def list_battles(lib_path: str, tracker_path: Optional[str] = None) -> list[dict]:
    """List prepared battles from the Battle Index with participant counts and warnings.

    Counts/validation parse each definition sheet's rows but do NOT resolve creature stats
    (cheap). If a tracker path is given, broken source-tab references are flagged by
    comparing names to the tracker's sheet names (no cell reads).
    """
    wb = _open(lib_path)
    if BATTLE_INDEX_SHEET not in wb.sheetnames:
        raise ValueError(f"Battle Library has no '{BATTLE_INDEX_SHEET}' sheet.")

    tracker_tabs: Optional[set] = None
    if tracker_path:
        try:
            tracker_tabs = set(_open(tracker_path).sheetnames)
        except Exception:
            tracker_tabs = None  # validation downgraded; not fatal

    battles: list[dict] = []
    for ir in _read_index_rows(wb[BATTLE_INDEX_SHEET]):
        battle = {
            "id": ir["battle_id"],
            "name": ir["name"] or ir["battle_id"],
            "status": ir["status"],
            "tags": [t.strip() for t in ir["tags"].split(",") if t.strip()],
            "notes": ir["notes"],
            "default_party_profile": ir["default_party_profile"],
            "definition_location": ir["definition_location"],
            "last_modified": ir["last_modified"],
            "counts": {"pcs": 0, "npcs": 0, "allies": 0, "hazards": 0, "events": 0, "groups": 0},
            "warnings": [],
        }
        loc = ir["definition_location"]
        if not loc:
            battle["warnings"].append("No definition location set.")
        elif loc not in wb.sheetnames:
            battle["warnings"].append(f"Definition sheet '{loc}' not found.")
        else:
            _count_and_validate(read_battle_list_rows(wb[loc]), battle, tracker_tabs)
        battles.append(battle)
    return battles


def load_battle(lib_path: str, tracker_path: str, battle_id: str) -> Encounter:
    """Resolve a battle definition into a self-contained live Encounter.

    Creature stats are pulled from Combat Tracker.xlsx at load time and snapshotted into
    the Encounter; later edits to source workbooks do not affect an in-progress battle.
    Raises ValueError on hard problems (unknown battle, missing definition sheet).
    """
    wb = _open(lib_path)
    if BATTLE_INDEX_SHEET not in wb.sheetnames:
        raise ValueError(f"Battle Library has no '{BATTLE_INDEX_SHEET}' sheet.")

    match = next((r for r in _read_index_rows(wb[BATTLE_INDEX_SHEET])
                  if r["battle_id"] == battle_id), None)
    if match is None:
        raise ValueError(f"Battle '{battle_id}' not found in Battle Index.")

    loc = match["definition_location"]
    if not loc:
        raise ValueError(f"Battle '{match['name'] or battle_id}' has no definition location.")
    if loc not in wb.sheetnames:
        raise ValueError(
            f"Battle '{match['name'] or battle_id}': definition sheet '{loc}' not found."
        )

    rows = read_battle_list_rows(wb[loc])
    resolve = make_combat_tracker_resolver(tracker_path)

    # source_path points at the Combat Tracker so Excel export can still write HP/status
    # back into the creature tabs.
    enc = Encounter(source_path=tracker_path)
    for row in rows:
        combatant = build_combatant_from_row(row, resolve)
        if combatant is None:
            continue
        enc.combatants[combatant.id] = combatant
        enc.order.append(combatant.id)
    return enc


# ── Builder support: list sources, read a definition, write/duplicate/delete ────

def list_sources(tracker_path: str) -> list[dict]:
    """Creature/NPC sources available in Combat Tracker for the builder's picker.

    A creature tab is any sheet (other than 'Battle List') with a name in B3.
    """
    wb = _open(tracker_path)
    out: list[dict] = []
    for name in wb.sheetnames:
        if name == "Battle List":
            continue
        ws = wb[name]
        disp = ws["B3"].value
        if not disp:
            continue
        out.append({
            "tab": name,
            "name": str(disp).strip(),
            "role": str(ws["B5"].value).strip() if ws["B5"].value else "",
            "ac": ws["B7"].value,
            "max_hp": ws["B8"].value,
        })
    out.sort(key=lambda s: s["name"].lower())
    return out


def get_battle_definition(lib_path: str, battle_id: str) -> dict:
    """Return a battle's metadata + roster rows for editing in the builder."""
    wb = _open(lib_path)
    if BATTLE_INDEX_SHEET not in wb.sheetnames:
        raise ValueError(f"Battle Library has no '{BATTLE_INDEX_SHEET}' sheet.")
    match = next((r for r in _read_index_rows(wb[BATTLE_INDEX_SHEET])
                  if r["battle_id"] == battle_id), None)
    if match is None:
        raise ValueError(f"Battle '{battle_id}' not found in Battle Index.")
    loc = match["definition_location"]
    rows = []
    if loc and loc in wb.sheetnames:
        rows = [r for r in read_battle_list_rows(wb[loc]) if r.get("type")]
    return {
        "id": match["battle_id"],
        "name": match["name"],
        "definition_location": loc,
        "status": match["status"],
        "tags": [t.strip() for t in match["tags"].split(",") if t.strip()],
        "notes": match["notes"],
        "default_party_profile": match["default_party_profile"],
        "rows": rows,
    }


def _next_battle_id(index_rows: list[dict]) -> str:
    nums = [int(m.group(1)) for r in index_rows
            if (m := re.match(r"BTL-(\d+)$", r["battle_id"]))]
    return f"BTL-{(max(nums) + 1) if nums else 1:03d}"


def _safe_sheet_name(name: str, wb, keep: Optional[str] = None) -> str:
    """A unique, Excel-legal sheet name derived from a battle name."""
    base = re.sub(_INVALID_SHEET_CHARS, " ", name).strip()[:28] or "Battle"
    taken = set(wb.sheetnames)
    if keep:
        taken.discard(keep)
    candidate, i = base, 2
    while candidate in taken:
        candidate = f"{base} {i}"[:31]
        i += 1
    return candidate


def _index_rownum(idx, battle_id: str) -> Optional[int]:
    for r in range(1, idx.max_row + 1):
        if str(idx.cell(row=r, column=1).value or "").strip() == battle_id:
            return r
    return None


def _ensure_index_header(idx) -> None:
    a1 = str(idx.cell(row=1, column=1).value or "").strip().lower()
    if a1 in ("battle id", "battle_id"):
        return
    has_data = any(idx.cell(row=1, column=c).value for c in range(1, len(_INDEX_HEADER) + 1))
    if has_data:
        idx.insert_rows(1)
    for col, h in enumerate(_INDEX_HEADER, start=1):
        idx.cell(row=1, column=col, value=h).font = Font(bold=True)


def _row_values(r: dict) -> list:
    qty = r.get("quantity")
    if qty in (None, ""):
        qty = 1
    out = []
    for key in ROSTER_COLUMNS:
        if key == "quantity":
            out.append(qty)
        elif key == "initiative_mode":
            out.append(r.get("initiative_mode") or "Individual")
        else:
            out.append(r.get(key, ""))
    return out


def save_battle(lib_path: str, battle: dict) -> dict:
    """Create or update a battle definition in the library (creating the file if needed).

    battle = {id?, name, status?, tags?(list|str), notes?, default_party_profile?,
              definition_location?, rows:[{type,name,npc_source,quantity,group_name,
              initiative_mode,hp_override,notes,starting_conditions,starting_status}, ...]}
    Returns {id, name, definition_location}.
    """
    name = (battle.get("name") or "").strip()
    if not name:
        raise ValueError("Battle name is required.")

    if os.path.exists(lib_path):
        wb = openpyxl.load_workbook(lib_path)
        if BATTLE_INDEX_SHEET not in wb.sheetnames:
            wb.create_sheet(BATTLE_INDEX_SHEET, 0)
    else:
        wb = openpyxl.Workbook()
        wb.active.title = BATTLE_INDEX_SHEET
    idx = wb[BATTLE_INDEX_SHEET]
    _ensure_index_header(idx)

    index_rows = _read_index_rows(idx)
    battle_id = (battle.get("id") or "").strip() or _next_battle_id(index_rows)
    existing = next((r for r in index_rows if r["battle_id"] == battle_id), None)

    loc = (battle.get("definition_location")
           or (existing["definition_location"] if existing else "")).strip()
    if not loc:
        loc = _safe_sheet_name(name, wb)

    # Replace the definition sheet contents wholesale.
    if loc in wb.sheetnames:
        del wb[loc]
    ws = wb.create_sheet(loc)
    ws.append(_DEF_HEADER)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in battle.get("rows", []):
        if not (r.get("type") or "").strip():
            continue
        ws.append(_row_values(r))

    tags = battle.get("tags", "")
    if isinstance(tags, list):
        tags = ",".join(tags)
    values = [
        battle_id, name, loc, battle.get("status") or "Draft", tags,
        battle.get("notes", ""), battle.get("default_party_profile", ""),
        date.today().isoformat(),
    ]
    rownum = _index_rownum(idx, battle_id)
    if rownum:
        for col, val in enumerate(values, start=1):
            idx.cell(row=rownum, column=col, value=val)
    else:
        idx.append(values)

    wb.save(lib_path)
    return {"id": battle_id, "name": name, "definition_location": loc}


def duplicate_battle(lib_path: str, battle_id: str) -> dict:
    """Copy a battle's setup into a new Draft battle (fresh id + sheet)."""
    defn = get_battle_definition(lib_path, battle_id)
    return save_battle(lib_path, {
        "name": f"{defn['name']} (copy)",
        "status": "Draft",
        "tags": defn["tags"],
        "notes": defn["notes"],
        "default_party_profile": defn["default_party_profile"],
        "rows": defn["rows"],
    })


def delete_battle(lib_path: str, battle_id: str) -> None:
    """Remove a battle's index row and its definition sheet."""
    wb = openpyxl.load_workbook(lib_path)
    if BATTLE_INDEX_SHEET not in wb.sheetnames:
        raise ValueError(f"Battle Library has no '{BATTLE_INDEX_SHEET}' sheet.")
    idx = wb[BATTLE_INDEX_SHEET]
    rownum = _index_rownum(idx, battle_id)
    if not rownum:
        raise ValueError(f"Battle '{battle_id}' not found in Battle Index.")
    loc = str(idx.cell(row=rownum, column=3).value or "").strip()
    idx.delete_rows(rownum, 1)
    if loc and loc in wb.sheetnames:
        del wb[loc]
    wb.save(lib_path)
