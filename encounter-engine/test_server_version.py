"""Check what code the server is actually running."""
import json, urllib.request

BASE = "http://localhost:5000"

def get(path):
    with urllib.request.urlopen(f"{BASE}{path}") as r:
        return r.read().decode()

def post(path, body):
    data = json.dumps(body).encode()
    req = urllib.request.Request(f"{BASE}{path}", data=data,
                                  headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r:
        return json.load(r)

# Load and fire an attack, check if attack_roll is in result
r = post("/api/load-excel", {"path": "C:/Users/jwallin/OneDrive/dnd - elemental crown/battle tracker/combat tracker.xlsx"})
enc = r["encounter"]
combatants = enc["combatants"]
korrum = next(c for c in combatants.values() if c["name"] == "Korrum")
rex = next(c for c in combatants.values() if c["name"] == "Rex")

# Try Veyrath's Burning Touch (to_hit=10, no save?) as a simpler case
veyrath = next(c for c in combatants.values() if c["name"] == "Veyrath")
print("Veyrath attacks:")
for a in veyrath["attacks"]:
    print(f"  {repr(a['name'])} to_hit={a.get('to_hit')} save_dc={a.get('save_dc')}")

r2 = post("/api/attack/roll", {
    "combatant_id": veyrath["id"],
    "attack_name": "Burning Touch",
    "target_id": rex["id"],
    "adv": False, "dis": False, "ac_override": None,
})
print("\nBurning Touch vs Rex (no AC):")
print("attack_roll present:", "attack_roll" in r2.get("result", {}))
print("save_dc present:", "save_dc" in r2.get("result", {}))
print(json.dumps(r2["result"], indent=2))
