"""Ask the running server which app.py file it loaded and what the attack route does."""
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

r = get("/api/debug-info")
print(json.dumps(r, indent=2))
