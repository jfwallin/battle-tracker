"""Find potential JS-breaking characters in combatant data."""
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
enc = state["encounter"]

# Look for backticks, backslashes, or other JS-dangerous chars in text fields
import re
danger = re.compile(r'[`\\]')

for cid, c in enc["combatants"].items():
    name = c.get("name","")
    effect_texts = []
    for atk in c.get("attacks", []):
        if atk.get("effect"):
            effect_texts.append(atk["effect"])

    for txt in [name] + effect_texts:
        if danger.search(txt):
            print(f"DANGER in {name}: {repr(txt[:100])}")

# Also check the notes/effect fields in traits/actions
for cid, c in enc["combatants"].items():
    name = c.get("name","")
    for section in ["traits","actions","bonus_actions","reactions","bloodied_effects"]:
        for item in c.get(section, []):
            for field in ["name","desc","description"]:
                txt = item.get(field,"")
                if danger.search(txt):
                    print(f"DANGER in {name}/{section}/{field}: {repr(txt[:100])}")

print("Check complete.")
