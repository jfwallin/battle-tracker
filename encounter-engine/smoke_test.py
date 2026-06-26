import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from engine.models import Encounter
from engine.excel_loader import load_encounter
from engine.combat import current_hp, derive_status, next_turn, roll_initiative, sort_initiative
from engine.dice import roll_attack, roll_damage, mob_hits
from engine.excel_exporter import export_encounter
import app as flask_app

print("All imports OK")

xlsx = os.path.join(os.path.dirname(__file__), "..", "combat tracker.xlsx")
enc = load_encounter(xlsx)
print(f"Loaded {len(enc.combatants)} combatants: {[c.name for c in enc.combatants.values()]}")

roll_initiative(enc, scope="all")
sort_initiative(enc)
print(f"Initiative order: {[enc.combatants[cid].name + ' ' + str(enc.combatants[cid].initiative) for cid in enc.order]}")

next_turn(enc)
print(f"After next_turn: round={enc.round} turn_index={enc.turn_index}")

c0 = enc.combatants[enc.order[0]]
print(f"Current HP of {c0.name}: {current_hp(c0, enc.log)} / {c0.max_hp}")
print(f"Status: {derive_status(c0, enc.log)}")

print("Smoke test PASSED")
