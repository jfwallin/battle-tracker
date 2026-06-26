"""Debug: print Korrum attack objects as returned by the API."""
import json, urllib.request

BASE = "http://localhost:5000"

def post(path, body):
    data = json.dumps(body).encode()
    req = urllib.request.Request(f"{BASE}{path}", data=data,
                                  headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r:
        return json.load(r)

r = post("/api/load-excel", {"path": "C:/Users/jwallin/OneDrive/dnd - elemental crown/battle tracker/combat tracker.xlsx"})
assert r["ok"], r

enc = r["encounter"]
combatants = enc["combatants"]
korrum = next(c for c in combatants.values() if c["name"] == "Korrum")
rex = next(c for c in combatants.values() if c["name"] == "Rex")

print("=== Korrum attacks from API ===")
for a in korrum["attacks"]:
    print(f"  name={repr(a['name'])} to_hit={repr(a.get('to_hit'))} save_dc={repr(a.get('save_dc'))}")

print(f"\nKorrum id={korrum['id']}, Rex id={rex['id']}")

print("\n=== Roll Slam vs Rex ===")
r2 = post("/api/attack/roll", {
    "combatant_id": korrum["id"],
    "attack_name": "Slam",
    "target_id": rex["id"],
    "adv": False,
    "dis": False,
    "ac_override": None,
})
print(json.dumps(r2, indent=2))
