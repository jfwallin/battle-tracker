"""Data containers for the encounter engine. No business logic."""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional
import uuid


@dataclass
class DamageDie:
    n: int
    die: int
    bonus: int
    damage_type: str  # "fire", "slashing", etc.; "" if unspecified

    def to_dict(self) -> dict:
        return {"n": self.n, "die": self.die, "bonus": self.bonus, "damage_type": self.damage_type}

    @classmethod
    def from_dict(cls, d: dict) -> "DamageDie":
        return cls(n=d["n"], die=d["die"], bonus=d["bonus"], damage_type=d.get("damage_type", ""))


@dataclass
class Attack:
    name: str
    attack_type: str           # "Melee", "Ranged", "Save", etc.
    to_hit: Optional[int]      # None if save-only or "—"
    reach: str                 # "5 ft." or "60/120 ft." etc.
    damage_dice: list[DamageDie] = field(default_factory=list)
    save_dc: Optional[int] = None
    save_ability: str = ""     # "STR", "DEX", etc.
    effect: str = ""

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "attack_type": self.attack_type,
            "to_hit": self.to_hit,
            "reach": self.reach,
            "damage_dice": [d.to_dict() for d in self.damage_dice],
            "save_dc": self.save_dc,
            "save_ability": self.save_ability,
            "effect": self.effect,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Attack":
        return cls(
            name=d["name"],
            attack_type=d.get("attack_type", ""),
            to_hit=d.get("to_hit"),
            reach=d.get("reach", ""),
            damage_dice=[DamageDie.from_dict(x) for x in d.get("damage_dice", [])],
            save_dc=d.get("save_dc"),
            save_ability=d.get("save_ability", ""),
            effect=d.get("effect", ""),
        )


@dataclass
class LimitedUse:
    name: str
    max_uses: int
    used: int = 0

    def to_dict(self) -> dict:
        return {"name": self.name, "max_uses": self.max_uses, "used": self.used}

    @classmethod
    def from_dict(cls, d: dict) -> "LimitedUse":
        return cls(name=d["name"], max_uses=d["max_uses"], used=d.get("used", 0))


@dataclass
class Member:
    id: str
    name: str
    max_hp: int
    conditions: list[str] = field(default_factory=list)
    can_attack: bool = True
    status_override: Optional[str] = None  # "Down", "Dead", "Fled", "Removed"

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "max_hp": self.max_hp,
            "conditions": self.conditions,
            "can_attack": self.can_attack,
            "status_override": self.status_override,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Member":
        return cls(
            id=d["id"],
            name=d["name"],
            max_hp=d["max_hp"],
            conditions=d.get("conditions", []),
            can_attack=d.get("can_attack", True),
            status_override=d.get("status_override"),
        )


@dataclass
class Combatant:
    id: str
    name: str
    combatant_type: str          # "PC", "NPC", "Ally", "Hazard", "Event"
    source_tab: Optional[str]    # creature sheet tab name, None for PCs/Hazards
    initiative: Optional[int]
    initiative_mod: int
    ac: Optional[int]
    max_hp: Optional[int]        # None for Hazards/Events
    temp_hp: int = 0
    conditions: list[str] = field(default_factory=list)
    condition_notes: dict[str, str] = field(default_factory=dict)  # condition -> duration note
    concentration: Optional[str] = None
    status_override: Optional[str] = None   # overrides derived status when set
    attacks: list[Attack] = field(default_factory=list)
    limited_use: list[LimitedUse] = field(default_factory=list)
    is_group: bool = False
    members: list[Member] = field(default_factory=list)
    initiative_mode: str = "Individual"  # "Individual|Shared|Grouped Attacks|Mob"
    attacks_per_round: int = 1
    reaction_used: bool = False
    notes: str = ""
    # Raw stat block strings for display
    speed: str = ""
    size_type_alignment: str = ""
    proficiency_bonus: int = 0
    ability_scores: dict[str, int] = field(default_factory=dict)  # STR/DEX/CON/INT/WIS/CHA
    ability_mods: dict[str, int] = field(default_factory=dict)
    save_bonuses: dict[str, int] = field(default_factory=dict)
    resistances: str = ""
    immunities: str = ""
    condition_immunities: str = ""
    senses: str = ""
    traits: list[dict] = field(default_factory=list)    # [{"name": ..., "desc": ...}]
    actions: list[dict] = field(default_factory=list)
    bonus_actions: list[dict] = field(default_factory=list)
    reactions: list[dict] = field(default_factory=list)
    bloodied_effects: list[dict] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "combatant_type": self.combatant_type,
            "source_tab": self.source_tab,
            "initiative": self.initiative,
            "initiative_mod": self.initiative_mod,
            "ac": self.ac,
            "max_hp": self.max_hp,
            "temp_hp": self.temp_hp,
            "conditions": self.conditions,
            "condition_notes": self.condition_notes,
            "concentration": self.concentration,
            "status_override": self.status_override,
            "attacks": [a.to_dict() for a in self.attacks],
            "limited_use": [lu.to_dict() for lu in self.limited_use],
            "is_group": self.is_group,
            "members": [m.to_dict() for m in self.members],
            "initiative_mode": self.initiative_mode,
            "attacks_per_round": self.attacks_per_round,
            "reaction_used": self.reaction_used,
            "notes": self.notes,
            "speed": self.speed,
            "size_type_alignment": self.size_type_alignment,
            "proficiency_bonus": self.proficiency_bonus,
            "ability_scores": self.ability_scores,
            "ability_mods": self.ability_mods,
            "save_bonuses": self.save_bonuses,
            "resistances": self.resistances,
            "immunities": self.immunities,
            "condition_immunities": self.condition_immunities,
            "senses": self.senses,
            "traits": self.traits,
            "actions": self.actions,
            "bonus_actions": self.bonus_actions,
            "reactions": self.reactions,
            "bloodied_effects": self.bloodied_effects,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Combatant":
        return cls(
            id=d["id"],
            name=d["name"],
            combatant_type=d["combatant_type"],
            source_tab=d.get("source_tab"),
            initiative=d.get("initiative"),
            initiative_mod=d.get("initiative_mod", 0),
            ac=d.get("ac"),
            max_hp=d.get("max_hp"),
            temp_hp=d.get("temp_hp", 0),
            conditions=d.get("conditions", []),
            condition_notes=d.get("condition_notes", {}),
            concentration=d.get("concentration"),
            status_override=d.get("status_override"),
            attacks=[Attack.from_dict(a) for a in d.get("attacks", [])],
            limited_use=[LimitedUse.from_dict(lu) for lu in d.get("limited_use", [])],
            is_group=d.get("is_group", False),
            members=[Member.from_dict(m) for m in d.get("members", [])],
            initiative_mode=d.get("initiative_mode", "Individual"),
            attacks_per_round=d.get("attacks_per_round", 1),
            reaction_used=d.get("reaction_used", False),
            notes=d.get("notes", ""),
            speed=d.get("speed", ""),
            size_type_alignment=d.get("size_type_alignment", ""),
            proficiency_bonus=d.get("proficiency_bonus", 0),
            ability_scores=d.get("ability_scores", {}),
            ability_mods=d.get("ability_mods", {}),
            save_bonuses=d.get("save_bonuses", {}),
            resistances=d.get("resistances", ""),
            immunities=d.get("immunities", ""),
            condition_immunities=d.get("condition_immunities", ""),
            senses=d.get("senses", ""),
            traits=d.get("traits", []),
            actions=d.get("actions", []),
            bonus_actions=d.get("bonus_actions", []),
            reactions=d.get("reactions", []),
            bloodied_effects=d.get("bloodied_effects", []),
        )


@dataclass
class LogEvent:
    id: str
    round: int
    turn: int
    source: str        # combatant name
    target: str        # combatant name or member name
    target_id: str     # combatant id
    member_id: Optional[str]  # member id if targeting a group member
    event_type: str    # "damage", "heal", "adjustment", "condition_add", "condition_remove", "note"
    amount: int        # HP delta (positive = damage or heal amount)
    damage_type: str
    attack_name: str
    notes: str
    ts: str            # ISO timestamp

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "round": self.round,
            "turn": self.turn,
            "source": self.source,
            "target": self.target,
            "target_id": self.target_id,
            "member_id": self.member_id,
            "event_type": self.event_type,
            "amount": self.amount,
            "damage_type": self.damage_type,
            "attack_name": self.attack_name,
            "notes": self.notes,
            "ts": self.ts,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "LogEvent":
        return cls(
            id=d["id"],
            round=d["round"],
            turn=d["turn"],
            source=d["source"],
            target=d["target"],
            target_id=d["target_id"],
            member_id=d.get("member_id"),
            event_type=d["event_type"],
            amount=d.get("amount", 0),
            damage_type=d.get("damage_type", ""),
            attack_name=d.get("attack_name", ""),
            notes=d.get("notes", ""),
            ts=d.get("ts", ""),
        )


@dataclass
class Encounter:
    round: int = 1
    turn_index: int = 0
    order: list[str] = field(default_factory=list)       # combatant ids in initiative order
    combatants: dict[str, Combatant] = field(default_factory=dict)
    log: list[LogEvent] = field(default_factory=list)
    source_path: str = ""
    started: bool = False

    def to_dict(self) -> dict:
        return {
            "round": self.round,
            "turn_index": self.turn_index,
            "order": self.order,
            "combatants": {k: v.to_dict() for k, v in self.combatants.items()},
            "log": [e.to_dict() for e in self.log],
            "source_path": self.source_path,
            "started": self.started,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Encounter":
        enc = cls(
            round=d.get("round", 1),
            turn_index=d.get("turn_index", 0),
            order=d.get("order", []),
            source_path=d.get("source_path", ""),
            started=d.get("started", False),
        )
        enc.combatants = {k: Combatant.from_dict(v) for k, v in d.get("combatants", {}).items()}
        enc.log = [LogEvent.from_dict(e) for e in d.get("log", [])]
        return enc


def make_id() -> str:
    return str(uuid.uuid4())[:8]
