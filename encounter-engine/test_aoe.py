"""Test AoE spell routes."""
import json, urllib.request

BASE = "http://localhost:5000"

def post(path, body=None):
    payload = json.dumps(body or {}).encode()
    req = urllib.request.Request(f"{BASE}{path}", data=payload,
                                  headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r:
        return json.load(r)

# Load encounter
post("/api/load-excel", {"path": "C:/Users/jwallin/OneDrive/dnd - elemental crown/battle tracker/combat tracker.xlsx"})
post("/api/initiative/sort", {})

# Get combatant IDs
import urllib.request as ur
with ur.urlopen(f"{BASE}/api/state") as r:
    state = json.load(r)
enc = state["encounter"]
combatants = enc["combatants"]
korrum = next(c for c in combatants.values() if c["name"] == "Korrum")
veyrath = next(c for c in combatants.values() if c["name"] == "Veyrath")
wisp = next(c for c in combatants.values() if c["name"] == "Wisp Pack")
rex = next(c for c in combatants.values() if c["name"] == "Rex")

print(f"Korrum immunities: {repr(korrum['immunities'])}")
print(f"Veyrath immunities: {repr(veyrath['immunities'])}")
print(f"Wisp Pack resistances: {repr(wisp['resistances'])}")

# Test roll-damage
r = post("/api/spell/roll-damage", {"expression": "8d6"})
print(f"\nRoll 8d6: {r}")

# Test AoE — Fireball (fire), DC 15 DEX, half on save
# Korrum: check if immune to fire; Veyrath: check; Wisp Pack: check
r2 = post("/api/spell/aoe", {
    "source": "Rex",
    "spell_name": "Fireball",
    "damage_roll": 28,
    "damage_type": "fire",
    "save_dc": 15,
    "save_ability": "DEX",
    "on_save": "half",
    "targets": [
        {"combatant_id": korrum["id"], "saved": False},
        {"combatant_id": veyrath["id"], "saved": True},
        {"combatant_id": wisp["id"], "saved": False},
    ]
})
print(f"\nAoE result ok={r2.get('ok')}")

# Check damage applied
with ur.urlopen(f"{BASE}/api/state") as r:
    state2 = json.load(r)
enc2 = state2["encounter"]
for name in ["Korrum", "Veyrath", "Wisp Pack"]:
    c = next(c for c in enc2["combatants"].values() if c["name"] == name)
    print(f"{name}: current_hp={c['current_hp']} total_damage_received={c['total_damage_received']}")

print("\nLast 3 log events:")
for e in enc2["log"][-3:]:
    print(f"  {e['source']} -> {e['target']}: {e['amount']} {e['damage_type']} ({e['attack_name']})")
