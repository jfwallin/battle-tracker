"""Test Burning Touch vs Rex after server restart."""
import json, urllib.request

BASE = "http://localhost:5000"

def post(path, body):
    data = json.dumps(body).encode()
    req = urllib.request.Request(f"{BASE}{path}", data=data,
                                  headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r:
        return json.load(r)

r = post("/api/load-excel", {"path": "C:/Users/jwallin/OneDrive/dnd - elemental crown/battle tracker/combat tracker.xlsx"})
enc = r["encounter"]
combatants = enc["combatants"]
veyrath = next(c for c in combatants.values() if c["name"] == "Veyrath")
rex = next(c for c in combatants.values() if c["name"] == "Rex")

print("Veyrath attacks from API:")
for a in veyrath["attacks"]:
    print(f"  name={repr(a['name'])} to_hit={a.get('to_hit')} save_dc={a.get('save_dc')}")

r2 = post("/api/attack/roll", {
    "combatant_id": veyrath["id"],
    "attack_name": "Burning Touch",
    "target_id": rex["id"],
    "adv": False, "dis": False, "ac_override": None,
})
print("\nAPI response:")
print(json.dumps(r2, indent=2))
