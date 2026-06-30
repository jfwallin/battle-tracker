"""Turn/round logic, HP derivation, status derivation."""
from __future__ import annotations
from datetime import datetime, timezone
from typing import Optional

from .models import Combatant, Encounter, LogEvent, Member, make_id
from .dice import roll as dice_roll


# ── HP derivation ─────────────────────────────────────────────────────────────

def current_hp_member(member: Member, combatant_id: str, log: list[LogEvent]) -> int:
    """Derive current HP for a group member from the log."""
    damage = sum(
        e.amount for e in log
        if e.target_id == combatant_id
        and e.member_id == member.id
        and e.event_type in ("damage", "adjustment_damage")
    )
    healing = sum(
        e.amount for e in log
        if e.target_id == combatant_id
        and e.member_id == member.id
        and e.event_type in ("heal", "adjustment_heal")
    )
    return member.max_hp - damage + healing


def current_hp(combatant: Combatant, log: list[LogEvent]) -> int:
    """Derive current HP for a combatant from the log. Groups sum member HPs."""
    if combatant.is_group:
        return sum(current_hp_member(m, combatant.id, log) for m in combatant.members)

    damage = sum(
        e.amount for e in log
        if e.target_id == combatant.id
        and e.member_id is None
        and e.event_type in ("damage", "adjustment_damage")
    )
    healing = sum(
        e.amount for e in log
        if e.target_id == combatant.id
        and e.member_id is None
        and e.event_type in ("heal", "adjustment_heal")
    )
    if combatant.max_hp is None:
        return 0
    return combatant.max_hp - damage + healing


def damage_received(combatant_id: str, log: list[LogEvent]) -> int:
    """Total damage logged against a combatant (used for PC count-up display)."""
    return sum(
        e.amount for e in log
        if e.target_id == combatant_id
        and e.member_id is None
        and e.event_type in ("damage", "adjustment_damage")
    )


def healing_received(combatant_id: str, log: list[LogEvent]) -> int:
    """Total healing logged for a combatant."""
    return sum(
        e.amount for e in log
        if e.target_id == combatant_id
        and e.member_id is None
        and e.event_type in ("heal", "adjustment_heal")
    )


def check_resistance(combatant: Combatant, damage_type: str) -> str:
    """Return 'immune', 'resistant', or 'normal' for a given damage type."""
    if not damage_type:
        return "normal"
    dt = damage_type.lower()
    if dt in combatant.immunities.lower():
        return "immune"
    if dt in combatant.resistances.lower():
        return "resistant"
    return "normal"


def derive_status(combatant: Combatant, log: list[LogEvent]) -> str:
    """Active/Bloodied/Down/Dead — unless DM has set a status_override."""
    if combatant.status_override:
        return combatant.status_override
    if combatant.max_hp is None:
        return "Active"  # Hazards/Events
    hp = current_hp(combatant, log)
    if hp > combatant.max_hp * 0.5:
        return "Active"
    if hp > 0:
        return "Bloodied"
    if hp <= 0:
        return "Down"
    return "Dead"


# ── Initiative ────────────────────────────────────────────────────────────────

def roll_initiative(encounter: Encounter, scope: str = "npcs") -> None:
    """
    Roll initiative for combatants based on scope.
    scope: "all" | "npcs" | "pcs"
    Shared/Grouped/Mob groups get one shared roll for the group.
    """
    already_rolled: set[str] = set()

    for cid in encounter.order:
        c = encounter.combatants[cid]
        if scope == "npcs" and c.combatant_type == "PC":
            continue
        if scope == "pcs" and c.combatant_type != "PC":
            continue

        if c.initiative_mode in ("Shared", "Grouped Attacks", "Mob"):
            # All members of this group share the same roll
            group_key = c.name
            if group_key in already_rolled:
                continue
            already_rolled.add(group_key)

        r = dice_roll(1, 20)
        c.initiative = r["total"] + c.initiative_mod


def set_initiative(encounter: Encounter, combatant_id: str, value: int) -> None:
    encounter.combatants[combatant_id].initiative = value


def sort_initiative(encounter: Encounter) -> None:
    """Sort encounter.order descending by initiative; stable tie-break preserves prior order."""
    def key(cid: str):
        c = encounter.combatants[cid]
        return c.initiative if c.initiative is not None else -999

    encounter.order.sort(key=key, reverse=True)
    # Reset turn to top after sort
    encounter.turn_index = 0


# ── Turn advancement ──────────────────────────────────────────────────────────

def _reset_turn_flags(combatant: Combatant) -> None:
    combatant.reaction_used = False
    combatant.attacks_used = 0
    combatant.bonus_action_used = False
    combatant.movement_used = False


def _reset_reactions_all(encounter: Encounter) -> None:
    """Reset reactions for every combatant at the start of a new round."""
    for c in encounter.combatants.values():
        c.reaction_used = False


def next_turn(encounter: Encounter) -> None:
    if not encounter.order:
        return
    encounter.turn_index += 1
    new_round = encounter.turn_index >= len(encounter.order)
    if new_round:
        encounter.turn_index = 0
        encounter.round += 1
        _reset_reactions_all(encounter)
    _reset_turn_flags(encounter.combatants[encounter.order[encounter.turn_index]])


def prev_turn(encounter: Encounter) -> None:
    if not encounter.order:
        return
    encounter.turn_index -= 1
    if encounter.turn_index < 0:
        if encounter.round > 1:
            encounter.round -= 1
        encounter.turn_index = len(encounter.order) - 1


def next_round(encounter: Encounter) -> None:
    encounter.round += 1
    encounter.turn_index = 0
    if encounter.order:
        _reset_turn_flags(encounter.combatants[encounter.order[0]])


def current_combatant(encounter: Encounter) -> Optional[Combatant]:
    if not encounter.order:
        return None
    return encounter.combatants[encounter.order[encounter.turn_index]]


# ── Damage & Healing ──────────────────────────────────────────────────────────

def _make_event(encounter: Encounter, **kwargs) -> LogEvent:
    c = current_combatant(encounter)
    return LogEvent(
        id=make_id(),
        round=encounter.round,
        turn=encounter.turn_index,
        source=kwargs.get("source", c.name if c else "DM"),
        target=kwargs["target"],
        target_id=kwargs["target_id"],
        member_id=kwargs.get("member_id"),
        event_type=kwargs["event_type"],
        amount=kwargs.get("amount", 0),
        damage_type=kwargs.get("damage_type", ""),
        attack_name=kwargs.get("attack_name", ""),
        notes=kwargs.get("notes", ""),
        ts=datetime.now(timezone.utc).isoformat(),
        outcome=kwargs.get("outcome", ""),
        roll=kwargs.get("roll"),
        hits=kwargs.get("hits"),
        crits=kwargs.get("crits", 0),
        attacks=kwargs.get("attacks"),
    )


def log_event(
    encounter: Encounter,
    event_type: str,
    target_id: str,
    *,
    source: str = "DM",
    target: Optional[str] = None,
    attack_name: str = "",
    outcome: str = "",
    roll: Optional[int] = None,
    damage_type: str = "",
    notes: str = "",
) -> LogEvent:
    """Append a non-HP narrative event (e.g. a miss or a save) for the battle report.
    amount stays 0 so HP derivation is unaffected."""
    tgt = encounter.combatants.get(target_id)
    event = _make_event(
        encounter,
        source=source,
        target=target if target is not None else (tgt.name if tgt else target_id),
        target_id=target_id,
        event_type=event_type,
        amount=0,
        damage_type=damage_type,
        attack_name=attack_name,
        outcome=outcome,
        roll=roll,
        notes=notes,
    )
    encounter.log.append(event)
    return event


def apply_damage(
    encounter: Encounter,
    target_id: str,
    amount: int,
    damage_type: str = "",
    attack_name: str = "",
    source: str = "DM",
    member_id: Optional[str] = None,
    outcome: str = "",
    roll: Optional[int] = None,
    hits: Optional[int] = None,
    crits: int = 0,
    attacks: Optional[int] = None,
) -> LogEvent:
    """Apply damage to a combatant (or group member). Temp HP depletes first.

    hits/crits/attacks summarize a group volley (so the battle report can show "5 hit of 8"
    instead of a single lumped damage event); leave them None for ordinary single hits."""
    target = encounter.combatants[target_id]

    # Handle temp HP (only for the whole combatant, not per-member)
    if member_id is None and target.temp_hp > 0:
        absorbed = min(target.temp_hp, amount)
        target.temp_hp -= absorbed
        amount -= absorbed
        if amount <= 0:
            # Fully absorbed — still log it
            event = _make_event(
                encounter,
                source=source,
                target=target.name,
                target_id=target_id,
                event_type="damage",
                amount=0,
                damage_type=damage_type,
                attack_name=attack_name,
                notes=f"Fully absorbed by temp HP",
                outcome=outcome,
                roll=roll,
                hits=hits,
                crits=crits,
                attacks=attacks,
            )
            encounter.log.append(event)
            return event

    member_name = target.name
    if member_id:
        m = next((m for m in target.members if m.id == member_id), None)
        if m:
            member_name = m.name

    event = _make_event(
        encounter,
        source=source,
        target=member_name if member_id else target.name,
        target_id=target_id,
        member_id=member_id,
        event_type="damage",
        amount=amount,
        damage_type=damage_type,
        attack_name=attack_name,
        outcome=outcome,
        roll=roll,
        hits=hits,
        crits=crits,
        attacks=attacks,
    )
    encounter.log.append(event)
    return event


def apply_heal(
    encounter: Encounter,
    target_id: str,
    amount: int,
    source: str = "DM",
    member_id: Optional[str] = None,
) -> LogEvent:
    target = encounter.combatants[target_id]
    member_name = None
    if member_id:
        m = next((m for m in target.members if m.id == member_id), None)
        if m:
            member_name = m.name

    event = _make_event(
        encounter,
        source=source,
        target=member_name or target.name,
        target_id=target_id,
        member_id=member_id,
        event_type="heal",
        amount=amount,
    )
    encounter.log.append(event)
    return event


def apply_temp_hp(encounter: Encounter, target_id: str, amount: int) -> None:
    """Temp HP: take the higher of current and new (don't stack)."""
    target = encounter.combatants[target_id]
    target.temp_hp = max(target.temp_hp, amount)


def undo_last(encounter: Encounter) -> Optional[LogEvent]:
    """Remove the last log event and return it."""
    if encounter.log:
        return encounter.log.pop()
    return None


def delete_log_event(encounter: Encounter, event_id: str) -> bool:
    for i, e in enumerate(encounter.log):
        if e.id == event_id:
            encounter.log.pop(i)
            return True
    return False


# ── Group damage allocation ───────────────────────────────────────────────────

def allocate_damage_focused(
    encounter: Encounter,
    target_id: str,
    member_id: str,
    amount: int,
    damage_type: str = "",
    attack_name: str = "",
    source: str = "DM",
) -> list[LogEvent]:
    """Apply all damage to one specific member."""
    return [apply_damage(
        encounter, target_id, amount,
        damage_type=damage_type, attack_name=attack_name,
        source=source, member_id=member_id,
    )]


def allocate_damage_frontload(
    encounter: Encounter,
    target_id: str,
    amount: int,
    damage_type: str = "",
    attack_name: str = "",
    source: str = "DM",
) -> list[LogEvent]:
    """
    Cascade damage to lowest-HP living member; overflow to next.
    Members with status_override Down/Dead/Fled/Removed are skipped.
    """
    target = encounter.combatants[target_id]
    living = [
        m for m in target.members
        if m.status_override not in ("Down", "Dead", "Fled", "Removed")
    ]
    # Sort by current HP ascending
    living.sort(key=lambda m: current_hp_member(m, target_id, encounter.log))

    events = []
    remaining = amount
    for m in living:
        if remaining <= 0:
            break
        hp = current_hp_member(m, target_id, encounter.log)
        dealt = min(remaining, max(hp, 1))  # always deal at least 1 to finish them
        evt = apply_damage(
            encounter, target_id, dealt,
            damage_type=damage_type, attack_name=attack_name,
            source=source, member_id=m.id,
        )
        events.append(evt)
        remaining -= dealt
    return events


def allocate_damage_even(
    encounter: Encounter,
    target_id: str,
    amount: int,
    damage_type: str = "",
    attack_name: str = "",
    source: str = "DM",
) -> list[LogEvent]:
    """Split damage evenly (floor) across living members."""
    target = encounter.combatants[target_id]
    living = [
        m for m in target.members
        if m.status_override not in ("Down", "Dead", "Fled", "Removed")
    ]
    if not living:
        return []
    per = amount // len(living)
    if per == 0:
        per = 1
    events = []
    for m in living:
        evt = apply_damage(
            encounter, target_id, per,
            damage_type=damage_type, attack_name=attack_name,
            source=source, member_id=m.id,
        )
        events.append(evt)
    return events


def allocate_damage_each(
    encounter: Encounter,
    target_id: str,
    amount: int,
    damage_type: str = "",
    attack_name: str = "",
    source: str = "DM",
) -> list[LogEvent]:
    """Apply full damage amount to every living member (AoE hits each one)."""
    target = encounter.combatants[target_id]
    living = [
        m for m in target.members
        if m.status_override not in ("Down", "Dead", "Fled", "Removed")
    ]
    if not living:
        return []
    events = []
    for m in living:
        evt = apply_damage(
            encounter, target_id, amount,
            damage_type=damage_type, attack_name=attack_name,
            source=source, member_id=m.id,
        )
        events.append(evt)
    return events


# ── Group heal allocation ─────────────────────────────────────────────────────

def allocate_heal_even(
    encounter: Encounter,
    target_id: str,
    amount: int,
    source: str = "DM",
) -> list[LogEvent]:
    """Split healing evenly (floor, min 1) across living members."""
    target = encounter.combatants[target_id]
    living = [
        m for m in target.members
        if m.status_override not in ("Dead", "Fled", "Removed")
    ]
    if not living:
        return []
    per = amount // len(living)
    if per == 0:
        per = 1
    events = []
    for m in living:
        events.append(apply_heal(encounter, target_id, per, source=source, member_id=m.id))
    return events


def allocate_heal_each(
    encounter: Encounter,
    target_id: str,
    amount: int,
    source: str = "DM",
) -> list[LogEvent]:
    """Apply the full heal amount to every living member."""
    target = encounter.combatants[target_id]
    living = [
        m for m in target.members
        if m.status_override not in ("Dead", "Fled", "Removed")
    ]
    if not living:
        return []
    events = []
    for m in living:
        events.append(apply_heal(encounter, target_id, amount, source=source, member_id=m.id))
    return events


# ── Conditions & limited-use ──────────────────────────────────────────────────

def add_condition(
    encounter: Encounter,
    combatant_id: str,
    condition: str,
    duration_note: str = "",
) -> None:
    c = encounter.combatants[combatant_id]
    if condition not in c.conditions:
        c.conditions.append(condition)
    c.condition_notes[condition] = duration_note
    encounter.log.append(_make_event(
        encounter,
        target=c.name,
        target_id=combatant_id,
        event_type="condition_add",
        notes=f"Added: {condition}" + (f" ({duration_note})" if duration_note else ""),
    ))


def remove_condition(encounter: Encounter, combatant_id: str, condition: str) -> None:
    c = encounter.combatants[combatant_id]
    if condition in c.conditions:
        c.conditions.remove(condition)
    c.condition_notes.pop(condition, None)
    encounter.log.append(_make_event(
        encounter,
        target=c.name,
        target_id=combatant_id,
        event_type="condition_remove",
        notes=f"Removed: {condition}",
    ))


def consume_limited_use(encounter: Encounter, combatant_id: str, ability_name: str) -> bool:
    c = encounter.combatants[combatant_id]
    for lu in c.limited_use:
        if lu.name == ability_name:
            if lu.used < lu.max_uses:
                lu.used += 1
                return True
    return False


def restore_limited_use(encounter: Encounter, combatant_id: str, ability_name: str) -> bool:
    c = encounter.combatants[combatant_id]
    for lu in c.limited_use:
        if lu.name == ability_name:
            if lu.used > 0:
                lu.used -= 1
                return True
    return False


# ── Attack pool (groups) ──────────────────────────────────────────────────────

def group_attack_pool(combatant: Combatant, log: list[LogEvent]) -> int:
    """Sum of attacks_per_round across living, can-attack group members."""
    if not combatant.is_group:
        return combatant.attacks_per_round
    return sum(
        combatant.attacks_per_round
        for m in combatant.members
        if m.can_attack
        and m.status_override not in ("Down", "Dead", "Fled", "Removed")
        and current_hp_member(m, combatant.id, log) > 0
    )
