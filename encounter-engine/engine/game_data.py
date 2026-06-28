"""Game Data (optional `Game_Data.xlsx`) — Stage 3.

Currently holds **party profiles**: reusable rosters of PCs/allies that can be attached to a
battle and imported into a live encounter, so battle definitions can stay enemies-only instead
of baking the party into each one. The app works fine without this file.

A single flat ``Party Profiles`` sheet keyed by Profile ID (one row per character) — easy to
hand-edit. Pure (no Flask) for unit testing.
"""
from __future__ import annotations
import os
from typing import Optional

import openpyxl
from openpyxl.styles import Font

from .models import Combatant, make_id

PARTY_SHEET = "Party Profiles"
PARTY_HEADER = [
    "Profile ID", "Character Name", "Type", "Default AC",
    "Default Max HP", "Initiative Modifier", "Notes",
]


def _open(path: str):
    return openpyxl.load_workbook(path, data_only=True)


def _to_int(v) -> Optional[int]:
    if v is None or v == "":
        return None
    try:
        return int(v)
    except (ValueError, TypeError):
        return None


def _cell(ws, r, c) -> str:
    v = ws.cell(row=r, column=c).value
    return str(v).strip() if v is not None else ""


def _header_row(ws) -> int:
    """Find the row whose column A is 'Profile ID' (else row 1)."""
    for row in ws.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        if cell.value and str(cell.value).strip().lower() in ("profile id", "profile_id"):
            return cell.row
    return 1


def _read_member_rows(ws) -> list[dict]:
    """All character rows (raw), each tagged with its profile id."""
    hr = _header_row(ws)
    members = []
    for r in range(hr + 1, ws.max_row + 1):
        pid = _cell(ws, r, 1)
        name = _cell(ws, r, 2)
        if not pid or not name:
            continue
        members.append({
            "profile_id": pid,
            "name": name,
            "type": _cell(ws, r, 3) or "PC",
            "ac": _to_int(ws.cell(row=r, column=4).value),
            "max_hp": _to_int(ws.cell(row=r, column=5).value),
            "init_mod": _to_int(ws.cell(row=r, column=6).value) or 0,
            "notes": _cell(ws, r, 7),
        })
    return members


def list_party_profiles(game_data_path: str) -> list[dict]:
    """Return profiles as ``[{id, members:[...], count}]`` (empty if no sheet)."""
    wb = _open(game_data_path)
    if PARTY_SHEET not in wb.sheetnames:
        return []
    by_id: dict[str, dict] = {}
    order: list[str] = []
    for m in _read_member_rows(wb[PARTY_SHEET]):
        pid = m["profile_id"]
        if pid not in by_id:
            by_id[pid] = {"id": pid, "members": []}
            order.append(pid)
        by_id[pid]["members"].append({k: m[k] for k in ("name", "type", "ac", "max_hp", "init_mod", "notes")})
    profiles = [by_id[p] for p in order]
    for p in profiles:
        p["count"] = len(p["members"])
    return profiles


def get_party_profile(game_data_path: str, profile_id: str) -> dict:
    profile = next((p for p in list_party_profiles(game_data_path) if p["id"] == profile_id), None)
    if profile is None:
        raise ValueError(f"Party profile '{profile_id}' not found.")
    return profile


def build_party_combatants(profile: dict) -> list[Combatant]:
    """Turn a party profile into Combatants (PCs/allies, no creature source)."""
    out = []
    for m in profile.get("members", []):
        out.append(Combatant(
            id=make_id(),
            name=m.get("name", ""),
            combatant_type=m.get("type") or "PC",
            source_tab=None,
            initiative=None,
            initiative_mod=m.get("init_mod") or 0,
            ac=m.get("ac"),
            max_hp=m.get("max_hp"),
            notes=m.get("notes", ""),
        ))
    return out


def apply_party_profile(encounter, game_data_path: str, profile_id: str) -> int:
    """Prepend a party profile's members to an encounter. Returns how many were added."""
    profile = get_party_profile(game_data_path, profile_id)
    party = build_party_combatants(profile)
    for c in party:
        encounter.combatants[c.id] = c
    # Party leads the roster (before the enemies already in order).
    encounter.order[:0] = [c.id for c in party]
    return len(party)


# ── Write side (editor) ─────────────────────────────────────────────────────────

def _member_values(profile_id: str, m: dict) -> list:
    return [
        profile_id, m.get("name", ""), m.get("type") or "PC",
        m.get("ac", ""), m.get("max_hp", ""), m.get("init_mod", 0), m.get("notes", ""),
    ]


def save_party_profile(game_data_path: str, profile: dict) -> dict:
    """Create or replace a profile's rows in Game_Data.xlsx (creating the file if needed).

    profile = {id, members:[{name,type,ac,max_hp,init_mod,notes}, ...]}.
    """
    pid = (profile.get("id") or "").strip()
    if not pid:
        raise ValueError("Profile ID is required.")

    if os.path.exists(game_data_path):
        wb = openpyxl.load_workbook(game_data_path)
        if PARTY_SHEET not in wb.sheetnames:
            wb.create_sheet(PARTY_SHEET, 0)
    else:
        wb = openpyxl.Workbook()
        wb.active.title = PARTY_SHEET
    ws = wb[PARTY_SHEET]

    # Read existing rows for OTHER profiles, then rewrite the whole sheet.
    existing = [m for m in _read_member_rows(ws) if m["profile_id"] != pid]
    for row in list(ws.iter_rows()):
        for cell in row:
            cell.value = None
    ws.append(PARTY_HEADER)
    for c in ws[1]:
        c.font = Font(bold=True)
    for m in existing:
        ws.append(_member_values(m["profile_id"], m))
    for m in profile.get("members", []):
        if not (m.get("name") or "").strip():
            continue
        ws.append(_member_values(pid, m))

    wb.save(game_data_path)
    return {"id": pid, "count": len([m for m in profile.get("members", []) if (m.get("name") or "").strip()])}


def delete_party_profile(game_data_path: str, profile_id: str) -> None:
    if not os.path.exists(game_data_path):
        raise ValueError("Game_Data.xlsx not found.")
    wb = openpyxl.load_workbook(game_data_path)
    if PARTY_SHEET not in wb.sheetnames:
        raise ValueError(f"No '{PARTY_SHEET}' sheet.")
    ws = wb[PARTY_SHEET]
    remaining = [m for m in _read_member_rows(ws) if m["profile_id"] != profile_id]
    for row in list(ws.iter_rows()):
        for cell in row:
            cell.value = None
    ws.append(PARTY_HEADER)
    for c in ws[1]:
        c.font = Font(bold=True)
    for m in remaining:
        ws.append(_member_values(m["profile_id"], m))
    wb.save(game_data_path)
