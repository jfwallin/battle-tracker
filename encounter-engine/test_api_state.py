"""Check /api/state after sort."""
import json, urllib.request

BASE = "http://localhost:5000"

def post(path, body=None):
    payload = json.dumps(body or {}).encode()
    req = urllib.request.Request(f"{BASE}{path}", data=payload,
                                  headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r:
        return json.load(r)

def get(path):
    with urllib.request.urlopen(f"{BASE}{path}") as r:
        return json.load(r)

post("/api/load-excel", {"path": "C:/Users/jwallin/OneDrive/dnd - elemental crown/battle tracker/combat tracker.xlsx"})
post("/api/initiative/sort", {})

state = get("/api/state")
enc = state.get("encounter", {})
print(f"order length: {len(enc.get('order', []))}")
print(f"combatants length: {len(enc.get('combatants', {}))}")
print(f"started: {enc.get('started')}")
print(f"current_combatant_id: {enc.get('current_combatant_id')}")
print(f"Keys in encounter: {list(enc.keys())}")
if enc.get('combatants'):
    first = next(iter(enc['combatants'].values()))
    print(f"\nFirst combatant keys: {list(first.keys())}")
    print(f"First combatant: name={first.get('name')}, type={first.get('combatant_type')}")
