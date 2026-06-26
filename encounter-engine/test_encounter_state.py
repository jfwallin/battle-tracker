"""Check encounter state after load + sort."""
import json, urllib.request

BASE = "http://localhost:5000"

def post(path, body):
    data = json.dumps(body).encode()
    req = urllib.request.Request(f"{BASE}{path}", data=data,
                                  headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r:
        return json.load(r)

def get(path):
    with urllib.request.urlopen(f"{BASE}{path}") as r:
        return json.load(r)

# Load excel
r = post("/api/load-excel", {"path": "C:/Users/jwallin/OneDrive/dnd - elemental crown/battle tracker/combat tracker.xlsx"})
assert r["ok"], r
print("Loaded. Combatants:")
for cid, c in r["encounter"]["combatants"].items():
    print(f"  {c['name']} ({c['combatant_type']}) init={c['initiative']}")

# Sort initiative
r2 = post("/api/initiative/sort", {})
print("\nAfter sort:")
print(f"  ok={r2.get('ok')} error={r2.get('error')}")
if r2.get("ok"):
    enc = r2["encounter"]
    print(f"  order={enc['order']}")
    print(f"  started={enc['started']}")
    for cid in enc["order"]:
        c = enc["combatants"][cid]
        print(f"  [{c['initiative']}] {c['name']} status={c.get('status')}")
else:
    print(json.dumps(r2, indent=2))
