"""Test the attack roll logic directly (no Flask) to confirm if/elif branch."""
import sys
sys.path.insert(0, '.')
from engine.excel_loader import load_encounter
from engine.dice import roll_attack, roll_damage
import random

enc = load_encounter('../combat tracker.xlsx')
korrum = next(c for c in enc.combatants.values() if c.name == 'Korrum')
rex = next(c for c in enc.combatants.values() if c.name == 'Rex')

atk = next(a for a in korrum.attacks if a.name == 'Slam')
print(f"atk.to_hit = {repr(atk.to_hit)}")
print(f"atk.save_dc = {repr(atk.save_dc)}")
print(f"atk.to_hit is not None: {atk.to_hit is not None}")

effective_ac = None  # Rex has no AC

result = {
    "combatant_id": korrum.id,
    "target_id": rex.id,
    "attack_name": atk.name,
    "target_ac": effective_ac,
}

adv = False
dis = False

if atk.to_hit is not None:
    print("\nTAKING: if atk.to_hit is not None branch")
    if effective_ac is not None:
        atk_roll = roll_attack(atk.to_hit, effective_ac, adv=adv, dis=dis)
    else:
        d1 = random.randint(1, 20)
        d2 = random.randint(1, 20) if adv or dis else None
        if adv and d2:
            nat = max(d1, d2)
        elif dis and d2:
            nat = min(d1, d2)
        else:
            nat = d1
        both = [d1, d2] if d2 else [d1]
        total = nat + atk.to_hit
        atk_roll = {
            "d20": nat, "both_dice": both, "total": total,
            "hit": None, "crit": nat == 20,
            "to_hit": atk.to_hit, "target_ac": None,
        }
    result["attack_roll"] = atk_roll
    if atk.damage_dice and (atk_roll.get("hit") or effective_ac is None or atk_roll.get("crit")):
        dmg = roll_damage(atk.damage_dice, crit=bool(atk_roll.get("crit")))
        result["damage"] = dmg
elif atk.save_dc is not None:
    print("\nTAKING: elif atk.save_dc branch  ← BUG if this prints!")
    result["save_dc"] = atk.save_dc

result["effect"] = atk.effect

import json
print("\nResult keys:", list(result.keys()))
print("attack_roll present:", "attack_roll" in result)
