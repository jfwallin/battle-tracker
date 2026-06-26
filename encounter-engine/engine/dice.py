"""Dice rolling engine for the encounter tracker."""
from __future__ import annotations
import random
from typing import Optional

from .models import DamageDie
from .excel_loader import parse_damage

# Mob attack table: required d20 roll -> attackers needed per successful hit
# Required roll: clamp(AC - toHit, 1, 20)
_MOB_TABLE: dict[int, int] = {
    1: 1, 2: 1, 3: 1, 4: 1, 5: 1,
    6: 2, 7: 2, 8: 2, 9: 2, 10: 2, 11: 2, 12: 2,
    13: 3, 14: 3,
    15: 4, 16: 4,
    17: 5, 18: 5,
    19: 10,
    20: 20,
}


def roll(n: int, die: int) -> dict:
    """Roll n dice of die sides. Returns individual rolls and total."""
    rolls = [random.randint(1, die) for _ in range(n)]
    return {"rolls": rolls, "total": sum(rolls)}


def roll_attack(
    to_hit: int,
    target_ac: int,
    adv: bool = False,
    dis: bool = False,
) -> dict:
    """
    Roll a single attack.
    Returns d20 (kept die), nat (natural value), total, hit, crit.
    Nat 20 always hits and crits; nat 1 always misses.
    Advantage/disadvantage each roll 2d20 and take high/low.
    """
    if adv and not dis:
        d1, d2 = random.randint(1, 20), random.randint(1, 20)
        nat = max(d1, d2)
        both = [d1, d2]
    elif dis and not adv:
        d1, d2 = random.randint(1, 20), random.randint(1, 20)
        nat = min(d1, d2)
        both = [d1, d2]
    else:
        nat = random.randint(1, 20)
        both = [nat]

    total = nat + to_hit
    crit = nat == 20
    if nat == 20:
        hit = True
    elif nat == 1:
        hit = False
    else:
        hit = total >= target_ac

    return {
        "d20": nat,
        "both_dice": both,
        "total": total,
        "hit": hit,
        "crit": crit,
        "to_hit": to_hit,
        "target_ac": target_ac,
    }


def roll_damage(damage_dice: list[DamageDie], crit: bool = False) -> dict:
    """
    Roll damage for a list of DamageDie.
    Crit doubles the number of dice (not the flat bonus).
    Returns per-clause dice, totals, and grand_total.
    """
    clause_results = []
    grand_total = 0
    for dd in damage_dice:
        n = dd.n * 2 if crit else dd.n
        r = roll(n, dd.die)
        clause_total = r["total"] + dd.bonus
        clause_results.append({
            "name": dd.damage_type,
            "n": n,
            "die": dd.die,
            "bonus": dd.bonus,
            "rolls": r["rolls"],
            "total": clause_total,
            "damage_type": dd.damage_type,
        })
        grand_total += clause_total
    return {
        "clauses": clause_results,
        "grand_total": grand_total,
        "crit": crit,
    }


def roll_damage_from_string(damage_str: str, crit: bool = False) -> dict:
    """Convenience: parse a damage string then roll it."""
    dice = parse_damage(damage_str)
    return roll_damage(dice, crit=crit)


def roll_batch(count: int, to_hit: int, target_ac: int) -> dict:
    """
    Roll count individual attacks. Returns all results, hit count, and damage dice list.
    Damage is NOT rolled here — caller decides which dice to roll after seeing hits.
    """
    results = [roll_attack(to_hit, target_ac) for _ in range(count)]
    hits = sum(1 for r in results if r["hit"])
    crits = sum(1 for r in results if r["crit"])
    return {
        "attacks": results,
        "count": count,
        "hits": hits,
        "crits": crits,
    }


def mob_hits(
    attackers: int,
    to_hit: int,
    target_ac: int,
    override: Optional[int] = None,
) -> dict:
    """
    Expected-hits calculation using the mob attack table.
    required_roll = clamp(AC - to_hit, 1, 20)
    per_hit = attackers needed per hit from table
    hits = floor(attackers / per_hit)
    override replaces the computed hit count if provided.
    """
    required = max(1, min(20, target_ac - to_hit))
    per_hit = _MOB_TABLE.get(required, 20)
    computed_hits = attackers // per_hit
    final_hits = override if override is not None else computed_hits
    return {
        "attackers": attackers,
        "required_roll": required,
        "per_hit": per_hit,
        "computed_hits": computed_hits,
        "hits": final_hits,
        "override_used": override is not None,
    }


def avg_damage(damage_dice: list[DamageDie]) -> float:
    """Average expected damage (no crit) for display purposes."""
    total = 0.0
    for dd in damage_dice:
        total += dd.n * (dd.die + 1) / 2.0 + dd.bonus
    return total
