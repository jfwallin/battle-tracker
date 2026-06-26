"""Check if the encounter page loads without errors."""
import json, urllib.request

BASE = "http://localhost:5000"

def post(path, body=None):
    payload = json.dumps(body or {}).encode()
    req = urllib.request.Request(f"{BASE}{path}", data=payload,
                                  headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r:
        return json.load(r)

def get_html(path):
    with urllib.request.urlopen(f"{BASE}{path}") as r:
        return r.read().decode('utf-8')

# Load + sort
post("/api/load-excel", {"path": "C:/Users/jwallin/OneDrive/dnd - elemental crown/battle tracker/combat tracker.xlsx"})
post("/api/initiative/sort", {})

# Fetch the encounter HTML page
html = get_html("/encounter")
# Check for error signs
if "Internal Server Error" in html or "Traceback" in html:
    print("SERVER ERROR in page!")
    print(html[:2000])
elif "combatant-list" in html:
    print("Page loaded OK. combatant-list div found.")
    # Check for JS syntax issues — look for the script block
    script_start = html.find('<script>')
    script_end = html.rfind('</script>')
    script_len = script_end - script_start if script_start > 0 else 0
    print(f"Script block size: {script_len} chars")
    # Check if ENC and render are in there
    print(f"'let ENC' in script: {'let ENC' in html}")
    print(f"'renderCombatants' in script: {'renderCombatants' in html}")
    print(f"'/api/state' in script: {'/api/state' in html}")
else:
    print("Unexpected page content:")
    print(html[:500])
