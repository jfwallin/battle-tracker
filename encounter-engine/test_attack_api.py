"""Test attack rolling against PC target (no AC) via the live API."""
import json
import urllib.request

BASE = "http://localhost:5000"

def post(path, body):
    data = json.dumps(body).encode()
    req = urllib.request.Request(f"{BASE}{path}", data=data,
                                  headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r:
        return json.load(r)

# Load encounter
r = post("/api/load-excel", {"path": "C:/Users/jwallin/OneDrive/dnd - elemental crown/battle tracker/combat tracker.xlsx"})
assert r["ok"], r

enc = r["encounter"]
combatants = enc["combatants"]

# Find Korrum (NPC) and Rex (PC)
korrum = next(c for c in combatants.values() if c["name"] == "Korrum")
rex = next(c for c in combatants.values() if c["name"] == "Rex")

print(f"Korrum id={korrum['id']} attacks={[a['name'] for a in korrum['attacks']]}")
print(f"Rex id={rex['id']} ac={rex['ac']} combatant_type={rex['combatant_type']}")

# Test 1: Roll Slam vs Rex with NO ac_override (AC unknown)
print("\n=== Test 1: Slam vs Rex, no AC ===")
r = post("/api/attack/roll", {
    "combatant_id": korrum["id"],
    "attack_name": "Slam",
    "target_id": rex["id"],
    "adv": False,
    "dis": False,
    "ac_override": None,
})
print(json.dumps(r, indent=2))

# Test 2: Roll Slam vs Rex with AC override = 16
print("\n=== Test 2: Slam vs Rex, AC override = 16 ===")
r2 = post("/api/attack/roll", {
    "combatant_id": korrum["id"],
    "attack_name": "Slam",
    "target_id": rex["id"],
    "adv": False,
    "dis": False,
    "ac_override": 16,
})
print(json.dumps(r2, indent=2))
