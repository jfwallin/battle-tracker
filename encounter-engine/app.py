"""Encounter Engine — Flask application."""
from __future__ import annotations
import json
import os
from datetime import datetime, timezone
from functools import wraps
from typing import Optional

from flask import Flask, jsonify, render_template, request, session

from engine.models import Encounter, LimitedUse, make_id
from engine.excel_loader import load_encounter
from engine.combat import (
    add_condition, allocate_damage_each, allocate_damage_even, allocate_damage_focused,
    allocate_damage_frontload, apply_damage, apply_heal, apply_temp_hp,
    check_resistance, consume_limited_use, current_combatant, current_hp,
    damage_received, delete_log_event, derive_status, group_attack_pool,
    healing_received, next_round, next_turn, prev_turn, remove_condition,
    restore_limited_use, roll_initiative, set_initiative, sort_initiative,
    undo_last,
)
from engine.dice import avg_damage, mob_hits, roll_attack, roll_batch, roll_damage

app = Flask(__name__)
app.secret_key = "encounter-engine-local"

STATE_PATH = os.path.join(os.path.dirname(__file__), "data", "encounter_state.json")

# Global in-memory encounter (single-user local)
_encounter: Optional[Encounter] = None


# ── Persistence ───────────────────────────────────────────────────────────────

def _save() -> None:
    if _encounter is not None:
        os.makedirs(os.path.dirname(STATE_PATH), exist_ok=True)
        with open(STATE_PATH, "w", encoding="utf-8") as f:
            json.dump(_encounter.to_dict(), f, indent=2)


def _load_state() -> Optional[Encounter]:
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, encoding="utf-8") as f:
            return Encounter.from_dict(json.load(f))
    return None


def _encounter_json() -> dict:
    """Return enriched encounter dict for the frontend."""
    if _encounter is None:
        return {}
    d = _encounter.to_dict()
    # Attach derived fields to each combatant
    for cid, cd in d["combatants"].items():
        c = _encounter.combatants[cid]
        cd["current_hp"] = current_hp(c, _encounter.log)
        cd["status"] = derive_status(c, _encounter.log)
        cd["attack_pool"] = group_attack_pool(c, _encounter.log)
        cd["total_damage_received"] = damage_received(c.id, _encounter.log)
        cd["total_healed"] = healing_received(c.id, _encounter.log)
        if c.is_group:
            for m_dict in cd["members"]:
                from engine.combat import current_hp_member
                mid = m_dict["id"]
                m = next((mm for mm in c.members if mm.id == mid), None)
                if m:
                    m_dict["current_hp"] = current_hp_member(m, cid, _encounter.log)
    # Attach current combatant id
    if _encounter.order:
        d["current_combatant_id"] = _encounter.order[_encounter.turn_index]
    else:
        d["current_combatant_id"] = None
    return d


def mutate(fn):
    """Decorator: run fn, save state, return updated encounter JSON."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        result = fn(*args, **kwargs)
        _save()
        if result is not None:
            return result
        return jsonify({"ok": True, "encounter": _encounter_json()})
    return wrapper


def require_encounter(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if _encounter is None:
            return jsonify({"error": "No encounter loaded"}), 400
        return fn(*args, **kwargs)
    return wrapper


# ── Pages ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    global _encounter
    saved = _load_state()
    return render_template(
        "setup.html",
        has_saved=saved is not None,
        saved_source=saved.source_path if saved else "",
    )


@app.route("/encounter")
def encounter_page():
    from flask import make_response
    if _encounter is None:
        return index()
    resp = make_response(render_template("encounter.html"))
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    return resp


# ── Setup / Load ──────────────────────────────────────────────────────────────

@app.route("/api/load-excel", methods=["POST"])
def api_load_excel():
    global _encounter
    data = request.get_json(force=True)
    path = data.get("path", "").strip()
    if not path or not os.path.exists(path):
        return jsonify({"error": f"File not found: {path}"}), 400
    try:
        _encounter = load_encounter(path)
        _save()
        return jsonify({"ok": True, "encounter": _encounter_json()})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/resume", methods=["POST"])
def api_resume():
    global _encounter
    saved = _load_state()
    if saved is None:
        return jsonify({"error": "No saved state"}), 404
    _encounter = saved
    return jsonify({"ok": True, "encounter": _encounter_json()})


@app.route("/api/state")
@require_encounter
def api_state():
    return jsonify({"encounter": _encounter_json()})


# ── Initiative ────────────────────────────────────────────────────────────────

@app.route("/api/initiative/roll", methods=["POST"])
@require_encounter
@mutate
def api_roll_initiative():
    data = request.get_json(force=True)
    scope = data.get("scope", "npcs")
    roll_initiative(_encounter, scope=scope)


@app.route("/api/initiative/set", methods=["POST"])
@require_encounter
@mutate
def api_set_initiative():
    data = request.get_json(force=True)
    set_initiative(_encounter, data["combatant_id"], int(data["value"]))


@app.route("/api/initiative/sort", methods=["POST"])
@require_encounter
@mutate
def api_sort_initiative():
    sort_initiative(_encounter)
    _encounter.started = True


# ── Turn ──────────────────────────────────────────────────────────────────────

@app.route("/api/turn/next", methods=["POST"])
@require_encounter
@mutate
def api_next_turn():
    next_turn(_encounter)


@app.route("/api/turn/prev", methods=["POST"])
@require_encounter
@mutate
def api_prev_turn():
    prev_turn(_encounter)


@app.route("/api/round/next", methods=["POST"])
@require_encounter
@mutate
def api_next_round():
    next_round(_encounter)


# ── HP ────────────────────────────────────────────────────────────────────────

@app.route("/api/combatant/<cid>/damage", methods=["POST"])
@require_encounter
@mutate
def api_damage(cid):
    data = request.get_json(force=True)
    apply_damage(
        _encounter, cid,
        amount=int(data["amount"]),
        damage_type=data.get("damage_type", ""),
        attack_name=data.get("attack_name", ""),
        source=data.get("source", "DM"),
        member_id=data.get("member_id"),
    )


@app.route("/api/combatant/<cid>/heal", methods=["POST"])
@require_encounter
@mutate
def api_heal(cid):
    data = request.get_json(force=True)
    apply_heal(
        _encounter, cid,
        amount=int(data["amount"]),
        source=data.get("source", "DM"),
        member_id=data.get("member_id"),
    )


@app.route("/api/combatant/<cid>/temp-hp", methods=["POST"])
@require_encounter
@mutate
def api_temp_hp(cid):
    data = request.get_json(force=True)
    apply_temp_hp(_encounter, cid, int(data["amount"]))


@app.route("/api/combatant/<cid>/status", methods=["POST"])
@require_encounter
@mutate
def api_set_status(cid):
    data = request.get_json(force=True)
    _encounter.combatants[cid].status_override = data.get("status") or None


@app.route("/api/combatant/<cid>/notes", methods=["POST"])
@require_encounter
@mutate
def api_set_notes(cid):
    data = request.get_json(force=True)
    _encounter.combatants[cid].notes = data.get("notes", "")


@app.route("/api/combatant/<cid>/member/<mid>/status", methods=["POST"])
@require_encounter
@mutate
def api_set_member_status(cid, mid):
    data = request.get_json(force=True)
    c = _encounter.combatants.get(cid)
    if not c:
        return jsonify({"error": "Unknown combatant"}), 400
    m = next((mm for mm in c.members if mm.id == mid), None)
    if not m:
        return jsonify({"error": "Unknown member"}), 400
    m.status_override = data.get("status") or None


@app.route("/api/combatant/<cid>/member/<mid>/can-attack", methods=["POST"])
@require_encounter
@mutate
def api_set_member_can_attack(cid, mid):
    data = request.get_json(force=True)
    c = _encounter.combatants.get(cid)
    if not c:
        return jsonify({"error": "Unknown combatant"}), 400
    m = next((mm for mm in c.members if mm.id == mid), None)
    if not m:
        return jsonify({"error": "Unknown member"}), 400
    m.can_attack = bool(data.get("can_attack", True))


# ── Attacks ───────────────────────────────────────────────────────────────────

@app.route("/api/attack/roll", methods=["POST"])
@require_encounter
def api_roll_attack():
    data = request.get_json(force=True)
    combatant_id = data["combatant_id"]
    attack_name = data["attack_name"]
    target_id = data["target_id"]
    adv = bool(data.get("adv", False))
    dis = bool(data.get("dis", False))
    ac_override = data.get("ac_override")  # DM-supplied AC for PC targets

    c = _encounter.combatants.get(combatant_id)
    target = _encounter.combatants.get(target_id)
    if not c or not target:
        return jsonify({"error": "Unknown combatant"}), 400

    atk = next((a for a in c.attacks if a.name == attack_name), None)
    if not atk:
        return jsonify({"error": f"Attack '{attack_name}' not found"}), 400

    # Use ac_override if provided (PC targets or DM correction), else fall back to stat block
    effective_ac = ac_override if ac_override is not None else target.ac

    result = {
        "combatant_id": combatant_id,
        "target_id": target_id,
        "attack_name": attack_name,
        "target_ac": effective_ac,
    }

    if atk.to_hit is not None:
        # Always roll the attack — even if AC unknown (effective_ac=None rolls without hit determination)
        if effective_ac is not None:
            atk_roll = roll_attack(atk.to_hit, effective_ac, adv=adv, dis=dis)
        else:
            # No AC: roll d20 + bonus, leave hit determination to DM
            from engine.dice import roll as dice_roll
            import random
            d1 = random.randint(1, 20)
            d2 = random.randint(1, 20) if adv or dis else None
            if adv and d2:
                nat = max(d1, d2)
            elif dis and d2:
                nat = min(d1, d2)
            else:
                nat = d1
            both = [d1, d2] if d2 else [d1]
            total = nat + atk.to_hit
            atk_roll = {
                "d20": nat, "both_dice": both, "total": total,
                "hit": None, "crit": nat == 20,
                "to_hit": atk.to_hit, "target_ac": None,
            }
        result["attack_roll"] = atk_roll
        # Roll damage if hit (or if AC unknown — always show damage so DM can apply)
        if atk.damage_dice and (atk_roll.get("hit") or effective_ac is None or atk_roll.get("crit")):
            dmg = roll_damage(atk.damage_dice, crit=bool(atk_roll.get("crit")))
            result["damage"] = dmg
            if atk.save_dc is not None:
                result["half_damage"] = dmg["grand_total"] // 2
        # Some attacks hit first then require a save (e.g., Slam) — include save info either way
        if atk.save_dc is not None:
            result["save_dc"] = atk.save_dc
            result["save_ability"] = atk.save_ability
    elif atk.save_dc is not None:
        result["save_dc"] = atk.save_dc
        result["save_ability"] = atk.save_ability
        if atk.damage_dice:
            dmg = roll_damage(atk.damage_dice)
            result["damage"] = dmg
            result["half_damage"] = dmg["grand_total"] // 2
    else:
        # No to_hit, no save — just a raw d20 for reference
        import random
        nat = random.randint(1, 20)
        result["attack_roll"] = {
            "d20": nat, "both_dice": [nat], "total": nat,
            "hit": None, "crit": nat == 20, "to_hit": 0, "target_ac": None,
        }

    result["effect"] = atk.effect
    result["attack_type"] = atk.attack_type
    return jsonify({"ok": True, "result": result})


@app.route("/api/attack/apply", methods=["POST"])
@require_encounter
@mutate
def api_apply_attack():
    data = request.get_json(force=True)
    target_id = data["target_id"]
    amount = int(data["amount"])
    apply_damage(
        _encounter, target_id,
        amount=amount,
        damage_type=data.get("damage_type", ""),
        attack_name=data.get("attack_name", ""),
        source=data.get("source", ""),
        member_id=data.get("member_id"),
    )


# ── Group damage allocation ───────────────────────────────────────────────────

@app.route("/api/group/<cid>/allocate", methods=["POST"])
@require_encounter
@mutate
def api_allocate(cid):
    data = request.get_json(force=True)
    mode = data.get("mode", "focused")
    amount = int(data["amount"])
    damage_type = data.get("damage_type", "")
    attack_name = data.get("attack_name", "")
    source = data.get("source", "DM")

    if mode == "focused":
        member_id = data.get("member_id")
        if not member_id:
            return jsonify({"error": "member_id required for focused mode"}), 400
        allocate_damage_focused(_encounter, cid, member_id, amount, damage_type, attack_name, source)
    elif mode == "frontload":
        allocate_damage_frontload(_encounter, cid, amount, damage_type, attack_name, source)
    elif mode == "even":
        allocate_damage_even(_encounter, cid, amount, damage_type, attack_name, source)
    else:
        return jsonify({"error": f"Unknown mode: {mode}"}), 400


# ── Group batch/mob attacks ───────────────────────────────────────────────────

@app.route("/api/group/<cid>/attack", methods=["POST"])
@require_encounter
def api_group_attack(cid):
    data = request.get_json(force=True)
    mode = data.get("mode", "batch")
    target_id = data["target_id"]
    attack_name = data["attack_name"]

    c = _encounter.combatants.get(cid)
    target = _encounter.combatants.get(target_id)
    if not c or not target:
        return jsonify({"error": "Unknown combatant"}), 400

    atk = next((a for a in c.attacks if a.name == attack_name), None)
    if not atk:
        return jsonify({"error": f"Attack '{attack_name}' not found"}), 400

    pool = group_attack_pool(c, _encounter.log)
    result = {"mode": mode, "pool": pool, "attack_name": attack_name}

    if mode == "batch":
        if atk.to_hit is None or target.ac is None:
            return jsonify({"error": "Batch requires to-hit and target AC"}), 400
        batch = roll_batch(pool, atk.to_hit, target.ac)
        if batch["hits"] > 0:
            dmg = roll_damage(atk.damage_dice * batch["hits"])
            batch["damage"] = dmg
        result["batch"] = batch

    elif mode == "mob":
        if atk.to_hit is None or target.ac is None:
            return jsonify({"error": "Mob requires to-hit and target AC"}), 400
        override = data.get("override")
        mob = mob_hits(pool, atk.to_hit, target.ac, override=override)
        if mob["hits"] > 0:
            dmg = roll_damage(atk.damage_dice * mob["hits"])
            mob["damage"] = dmg
        result["mob"] = mob

    elif mode == "average":
        avg = avg_damage(atk.damage_dice)
        hits = data.get("hits", pool)
        result["average"] = {"avg_per_hit": avg, "hits": hits, "total": avg * hits}

    return jsonify({"ok": True, "result": result})


# ── AoE Spells ────────────────────────────────────────────────────────────────

@app.route("/api/spell/roll-damage", methods=["POST"])
@require_encounter
def api_spell_roll_damage():
    """Roll a dice expression (e.g. '8d6') and return the result. No state mutation."""
    import re as _re
    data = request.get_json(force=True)
    expr = (data.get("expression") or "").strip().lower().replace(" ", "")
    m = _re.match(r"^(\d+)d(\d+)([+-]\d+)?$", expr)
    if not m:
        return jsonify({"error": f"Invalid expression: {expr!r}. Use format like '8d6' or '2d8+3'"}), 400
    n, die = int(m.group(1)), int(m.group(2))
    bonus = int(m.group(3) or 0)
    from engine.dice import roll as dice_roll
    result = dice_roll(n, die)
    total = result["total"] + bonus
    return jsonify({"ok": True, "rolls": result["rolls"], "bonus": bonus, "total": total})


@app.route("/api/spell/aoe", methods=["POST"])
@require_encounter
@mutate
def api_spell_aoe():
    """Apply AoE spell damage to multiple targets, respecting resistance/immunity and saves."""
    data = request.get_json(force=True)
    print(f"[AoE] received: {data}", flush=True)
    source = data.get("source", "DM")
    spell_name = data.get("spell_name", "Spell")
    damage_roll = int(data["damage_roll"])
    damage_type = data.get("damage_type", "")
    on_save = data.get("on_save", "half")   # "half" or "none"
    targets = data.get("targets", [])
    print(f"[AoE] dmg={damage_roll} type={damage_type} on_save={on_save} targets={targets}", flush=True)

    for t in targets:
        cid = t["combatant_id"]
        saved = bool(t.get("saved", False))
        combatant = _encounter.combatants.get(cid)
        if not combatant:
            continue

        # DM override wins; otherwise auto-detect from combatant strings
        res_override = t.get("res_override")
        res = res_override if res_override in ("immune", "resistant", "normal") else check_resistance(combatant, damage_type)
        if res == "immune":
            continue  # no damage, skip entirely

        # Apply save result
        if saved:
            base = damage_roll // 2 if on_save == "half" else 0
        else:
            base = damage_roll

        # Apply resistance
        final_dmg = base // 2 if res == "resistant" else base

        if final_dmg <= 0:
            continue

        print(f"[AoE] applying {final_dmg} to {combatant.name} (res={res}, saved={saved})", flush=True)
        if combatant.is_group:
            group_mode = t.get("group_mode", "split")
            if group_mode == "each":
                allocate_damage_each(_encounter, cid, final_dmg, damage_type, spell_name, source)
            else:
                allocate_damage_even(_encounter, cid, final_dmg, damage_type, spell_name, source)
        else:
            apply_damage(_encounter, cid, final_dmg, damage_type, spell_name, source)


# ── Conditions ────────────────────────────────────────────────────────────────

@app.route("/api/condition", methods=["POST"])
@require_encounter
@mutate
def api_condition():
    data = request.get_json(force=True)
    cid = data["combatant_id"]
    if "add" in data:
        add_condition(_encounter, cid, data["add"], data.get("duration_note", ""))
    if "remove" in data:
        remove_condition(_encounter, cid, data["remove"])


# ── Limited use ───────────────────────────────────────────────────────────────

@app.route("/api/limited-use/consume", methods=["POST"])
@require_encounter
@mutate
def api_consume_limited_use():
    data = request.get_json(force=True)
    consume_limited_use(_encounter, data["combatant_id"], data["ability_name"])


@app.route("/api/limited-use/restore", methods=["POST"])
@require_encounter
@mutate
def api_restore_limited_use():
    data = request.get_json(force=True)
    restore_limited_use(_encounter, data["combatant_id"], data["ability_name"])


@app.route("/api/limited-use/add", methods=["POST"])
@require_encounter
@mutate
def api_add_limited_use():
    data = request.get_json(force=True)
    cid = data["combatant_id"]
    c = _encounter.combatants[cid]
    c.limited_use.append(LimitedUse(
        name=data["name"],
        max_uses=int(data["max_uses"]),
        used=0,
    ))


# ── Log ───────────────────────────────────────────────────────────────────────

@app.route("/api/log/undo", methods=["POST"])
@require_encounter
@mutate
def api_undo():
    undo_last(_encounter)


@app.route("/api/log/<event_id>", methods=["DELETE"])
@require_encounter
@mutate
def api_delete_log(event_id):
    delete_log_event(_encounter, event_id)


@app.route("/api/log/<event_id>", methods=["PATCH"])
@require_encounter
@mutate
def api_edit_log(event_id):
    data = request.get_json(force=True)
    for e in _encounter.log:
        if e.id == event_id:
            if "amount" in data:
                e.amount = int(data["amount"])
            if "notes" in data:
                e.notes = data["notes"]
            break


# ── Export ────────────────────────────────────────────────────────────────────

@app.route("/api/export-excel", methods=["POST"])
@require_encounter
def api_export_excel():
    data = request.get_json(force=True)
    out_path = data.get("path", "")
    if not out_path:
        return jsonify({"error": "path required"}), 400
    try:
        from engine.excel_exporter import export_encounter
        export_encounter(_encounter, out_path)
        return jsonify({"ok": True, "path": out_path})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/debug-info")
def api_debug_info():
    import sys
    template_path = os.path.join(app.template_folder, "encounter.html")
    template_mtime = os.path.getmtime(template_path) if os.path.exists(template_path) else None
    return jsonify({
        "file": __file__,
        "python": sys.executable,
        "version": "v2-attack-fix",
        "template_path": template_path,
        "template_mtime": template_mtime,
        "cwd": os.getcwd(),
    })


if __name__ == "__main__":
    import subprocess
    # Kill any process holding port 5000 (including stale reloader children)
    try:
        result = subprocess.run(["netstat", "-ano"], capture_output=True, text=True)
        for line in result.stdout.splitlines():
            if ":5000" in line and "LISTENING" in line:
                parts = line.split()
                pid = int(parts[-1])
                if pid != os.getpid():
                    print(f"[startup] Killing PID {pid} on port 5000", flush=True)
                    subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
    except Exception as e:
        print(f"[startup] Port cleanup skipped: {e}", flush=True)

    # use_reloader=False prevents Flask from spawning a child watcher process
    # that holds the port and survives after the parent is killed.
    # debug=True is kept for useful error pages.
    app.run(debug=True, port=5000, use_reloader=False)
