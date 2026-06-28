# Encounter Engine — Design & Architecture (current build)

This is the living design document for the app **as built**. It supersedes the original
aspirational spec in [`../design-spec.md`](../design-spec.md) (kept for history) and is the
place to look for how the app actually works today. Companion: the
[Monster Card Authoring Guide](MONSTER_CARD_TEMPLATE.md) for the Excel input contract.

---

## 1. Purpose

A local, single-user web app that runs a D&D 5e combat encounter loaded from an Excel
workbook. It replaces a formula spreadsheet that couldn't roll dice, apply damage to targets,
or track turns. Design priority: **speed and low bookkeeping during live play**. The app
assists DM judgment; it does not auto-adjudicate full 5e rules.

## 2. Tech stack

- **Backend:** Python 3.10+, Flask. `openpyxl` for Excel read/write.
- **Frontend:** Server-rendered Jinja2 templates + vanilla JS (no framework). Most page logic
  lives in inline `<script>` blocks in the templates; `static/app.js` is an intentional no-op
  placeholder.
- **State:** one in-memory `Encounter` object, autosaved to `data/encounter_state.json` after
  every mutation. Excel is read on load and written only on explicit export.
- **Run mode:** `python app.py`, localhost:5000, single user, no auth.

## 3. File structure

```
encounter-engine/
  app.py                    # Flask app: all routes, persistence, the @mutate/@require_encounter decorators
  engine/
    models.py               # dataclasses: Encounter, Combatant, Member, Attack, DamageDie, LimitedUse, LogEvent
    excel_loader.py         # workbook -> Encounter; also the damage/to-hit/save string parsers
    excel_exporter.py       # Encounter -> workbook copy (explicit export only)
    dice.py                 # dice parser + roller: roll_attack, roll_damage, roll_batch, mob_hits, avg_damage
    combat.py               # HP/status derivation, turn/round logic, damage/heal, group allocation, conditions
    workspace.py            # workspace directory discovery + remembered-workspace config (Stage 1)
    battle_library.py       # list/load/build reusable battle definitions referencing Combat Tracker (Stage 1–2)
    game_data.py            # optional Game_Data.xlsx: party profiles (Stage 3)
  templates/
    base.html               # layout shell
    setup.html              # workspace + battle picker + load/resume + roster preview + initiative + Start
    encounter.html          # the live combat dashboard (cards + all modals + most JS)
  static/  (style.css, app.js, dice_roller.js)
  data/    (encounter_state.json, workspace.json)
  tests/   (pytest unit tests for loader, dice, combat, workspace, battle_library; make_sample_library.py)
  DESIGN.md, MONSTER_CARD_TEMPLATE.md, Encounter Engine Expansion.md
```

> Root-level `test_*.py` scripts hit a live server and are ad-hoc debugging tools, **not** the
> unit suite. Run `python -m pytest tests/` for the real tests.

## 4. Runtime data model

Persisted JSON mirrors the dataclasses in `engine/models.py`.

```
Encounter { round, turn_index, order:[combatantId...], combatants:{id:Combatant}, log:[LogEvent], source_path, started }

Combatant {
  id, name, combatant_type,         # PC | NPC | Ally | Hazard | Event
  source_tab, initiative, initiative_mod, ac, max_hp, temp_hp,
  conditions:[str], condition_notes:{cond:note}, concentration, status_override,
  attacks:[Attack], limited_use:[LimitedUse],
  is_group, members:[Member], initiative_mode, attacks_per_round, reaction_used, notes,
  # display-only stat block: speed, size_type_alignment, proficiency_bonus,
  #   ability_scores/ability_mods/save_bonuses {STR..CHA}, resistances, immunities,
  #   condition_immunities, senses, traits, actions, bonus_actions, reactions, bloodied_effects
}

Member { id, name, max_hp, conditions, can_attack, status_override }   # group member
Attack { name, attack_type, to_hit:int|null, reach, damage_dice:[DamageDie], save_dc:int|null, save_ability, effect }
DamageDie { n, die, bonus, damage_type }
LimitedUse { name, max_uses, used }
LogEvent { id, round, turn, source, target, target_id, member_id, event_type, amount, damage_type, attack_name, notes, ts }
```

**Derived (enriched) fields** added by `_encounter_json()` for the frontend, not stored:
`current_hp`, `status`, `attack_pool`, `total_damage_received`, `total_healed` per combatant;
`current_hp` per member; and top-level `current_combatant_id`.

### HP is log-derived (the key invariant)

There is **no stored `current_hp`**. A combatant's HP is always recomputed from the log:
`max_hp − Σ(damage events) + Σ(heal events)` (filtered by `target_id`, and by `member_id` for
group members). The log is the single source of truth, so undo/edit/delete of any log row
recomputes HP for free. See `current_hp` / `current_hp_member` in `engine/combat.py`.

- **Temp HP** is stored on the combatant and depleted first on damage (whole-combatant only,
  not per member). "Take the higher" semantics — temp HP does not stack.
- **HP is not capped at max** — healing can exceed `max_hp` (by design; PCs are tracked by a
  damage/heal count-up rather than a hard HP bar — see §10).
- **Status** derives from HP: `> 50%` Active, `≤ 50%` Bloodied, `≤ 0` Down — unless the DM sets
  a `status_override` (Fled/Removed/etc.). Hazards/Events (no `max_hp`) are always Active.

## 5. Core mechanics

### Initiative & turns
- "Roll NPC/PC/All Initiative" → d20 + `initiative_mod`; Shared/Grouped/Mob groups get **one**
  shared roll. Results are stored static values (no re-roll on unrelated actions).
- Manual inline initiative entry on both the roster and encounter pages.
- Sort descending, stable tie-break. `Next/Prev Turn` advances `turn_index`; wrapping past the
  end auto-increments `round`. Turn start resets per-turn flags (`reaction_used`).

### Damage & healing
- Quick damage/heal inputs on every card write a `LogEvent`; HP recomputes instantly.
- Damage/heal amount fields accept **`+`/`-` sum expressions** (e.g. `12+20` → 32) via
  `parseAmount()` — handy for applying several attack rolls at once. Inputs are `type="text"`
  so the operators can be typed.
- **Healing targets** = any combatant except Hazard/Event, **including self** (`healTargetsFor`).
  PCs often have no `max_hp`, so heal targeting deliberately does **not** filter on it.
- A single-target heal of a **group** is distributed across members (else it'd be lost, since
  group HP is summed from per-member events).

### Group damage allocation (`engine/combat.py`)
- **Focused** (one member), **Frontload** (cascade into lowest-HP living member, overflow to
  next), **Even** (split, floor, min 1), **Each** (full amount to every living member, for AoE).

### Attacks
- Each attack renders as a button → modal: pick target, adv/dis, roll to-hit vs AC (nat 20
  hits+crits, nat 1 misses), roll crit-aware damage, "Apply". A DM-supplied AC override is
  honored (required for PC targets with no stat AC). Save-based attacks show DC + ability for
  per-target pass/fail and full/half application.
- **Crit rule:** doubles the **dice**, not the flat bonus (`roll_damage(crit=True)`).

### Group attacks (`/api/group/<id>/attack`)
Three modes, explained inline in the modal:
- **Batch** — rolls a separate d20 per attacker in the pool, then rolls damage **per hit**, so
  each crit doubles **its own** dice. Best for ≤ ~10 attackers.
- **Mob** — no d20s; estimates hits from the 5e mob table (`required = clamp(AC − toHit, 1, 20)`,
  attackers-per-hit lookup), then rolls damage once per estimated hit. Fast for swarms; no crits.
- **Average** — hits × average damage, fully deterministic.
- **Attack pool** = Σ `attacks_per_round` over members that are alive, `can_attack`, and not
  Down/Dead/Fled/Removed.

### AoE spells (`/api/spell/aoe`, the AoE modal)
- Roll a dice expression or type a total; pick a damage type; select targets.
- Per-target **resistance/immunity** auto-detected from the stat block (substring match) with a
  click-to-cycle override (normal/resistant/immune).
- Per-target **save** workflow: a roll button (uses the creature's save bonus for the ability),
  plus **advantage/disadvantage** (cycle button) and an **extra ±modifier** field — because each
  PC/NPC has different save mods (and PCs have none stored, so the field is where you enter it).
- **On-save outcome:** Half / None, or **No save (auto-hit)** which hides per-target save
  adjudication and applies full damage. Group targets choose Split ÷ or Each ✕.

### Mass healing (`/api/spell/mass-heal`, the Mass Heal modal)
- Mirrors the AoE target-picker for healing: roll/enter an amount, select targets (self
  included), Heal All. Group targets use Split (`allocate_heal_even`) or Each (`allocate_heal_each`).

### Conditions, limited-use, hazards
- Multi-select condition chips + free-text duration note per combatant (no rule enforcement).
- Limited-use counters (legendary resistance, recharge, slots) — click to consume/restore; DM
  can add ad-hoc counters.
- Hazards/Events are non-creature rows: initiative + notes, no HP.

### Manual stat edits
- On the roster page, **AC and Max HP are editable** (`/api/combatant/<id>/stats`) — e.g. to
  beef up an NPC. For a group, Max HP is per-member and syncs to all members. **These edits
  live only in the encounter state and are never written back to Excel.**

### Party Profiles (Stage 3 of the expansion, party-profiles slice)

- Optional `Game_Data.xlsx` holds **party profiles** — reusable PC/ally rosters
  (`engine/game_data.py`: a flat `Party Profiles` sheet keyed by Profile ID;
  `list/get/save/delete_party_profile`, `build_party_combatants`, `apply_party_profile`). The
  app runs fine without the file.
- On `POST /api/battles/<id>/load`, a `party_profile` body value attaches a party:
  a profile id, `""` (none), or `"__default__"`/absent (use the battle's **Default Party
  Profile** column). The party's members are prepended to the encounter roster. This lets
  battles be authored enemies-only and get the party attached at load time, instead of baking
  PCs into every definition.
- Setup UI: a **Party Profiles** card (list + New/Edit/Delete + a member editor:
  name/type/AC/HP/init/notes) and a per-battle **Party** picker on each battle card. Routes:
  `/api/party-profiles[/<id>]` (GET/POST save/DELETE). PCs in a profile have no creature source —
  AC/HP are DM-entered (or left blank, tracked by the usual damage/heal count-up).

### Battle Builder (Stage 2 of the expansion)

- Battles can be created/edited/duplicated/deleted **in-app** without hand-editing Excel
  (`engine/battle_library.py`: `list_sources`, `get_battle_definition`, `save_battle`,
  `duplicate_battle`, `delete_battle`; routes `/api/sources`, `/api/battles/<id>/definition`,
  `/api/battles/save`, `/api/battles/<id>/duplicate`, DELETE `/api/battles/<id>`). The setup
  page has a builder panel: name/status/tags/notes, a searchable Combat Tracker source picker,
  and an editable roster table (type/qty/group/init-mode/HP/status/notes) plus add-PC/Hazard/Event.
- `save_battle` creates `Battle_Library.xlsx` (with a `Battle Index`) if it doesn't exist yet, so
  the first battle can be authored from an empty workspace; editing reuses the same definition
  sheet, creating-vs-updating keyed by Battle ID (`BTL-NNN`, auto-numbered). Saving fails clean
  (HTTP 423) if the workbook is open in Excel.

### Workspace & Battle Library (Stage 1 of the expansion)

- A **workspace** is a directory of standard files, remembered in `data/workspace.json`
  (`engine/workspace.py`): `Combat Tracker.xlsx` (required), `Battle_Library.xlsx` (optional,
  enables prepared battles), `Game_Data.xlsx` (optional, not used yet). Filename matching is
  **case-insensitive** (the real library file is `combat tracker.xlsx`). Discovery reports each
  file found/missing with a specific message; only the Combat Tracker is required.
- **`Battle_Library.xlsx`** holds reusable battle definitions (`engine/battle_library.py`): a
  `Battle Index` sheet (Battle ID, Name, Definition Location, Status, Tags, Notes, Default Party
  Profile, Last Modified) plus **one definition sheet per battle, shaped exactly like the
  `Battle List`** (columns A–H), optionally with cols I–L (Starting Conditions, Starting Status,
  Variant ID, Trigger). `NPC Source` cells reference creature tabs in `Combat Tracker.xlsx`.
- **Cross-workbook resolution** is the key enabler: `excel_loader` was refactored into reusable
  units — `read_battle_list_rows(ws)`, `build_combatant_from_row(row, resolve)`, and resolver
  factories `make_workbook_resolver(wb)` / `make_combat_tracker_resolver(path)`. The roster can
  come from a battle definition while creature stats resolve from the Combat Tracker. Legacy
  `load_encounter(path)` is now a thin wrapper over these (behavior/API unchanged).
- **Loading a battle** (`load_battle`) builds a normal `Encounter` snapshot (`source_path` =
  the tracker, so Excel export still works) — later edits to source workbooks never alter an
  in-progress battle. The setup page's battle picker reuses the existing roster-preview →
  initiative → Start flow unchanged. The legacy single-file load remains for the no-workspace
  case. Validation: hard errors for missing files/sheets; soft `warnings[]` for broken source
  tabs, bad quantities/modes, and not-yet-supported Variant/Trigger columns.

## 6. REST API

All mutating routes return the updated, enriched encounter JSON (`{ok, encounter}`) so the UI
re-renders. Pages: `GET /`, `GET /encounter`.

| Method | Route | Purpose |
|--------|-------|---------|
| POST | `/api/load-excel` | import roster + stat blocks from a workbook path (single-file path) |
| GET  | `/api/workspace` · POST `/api/workspace/set` `{dir}` | read / set the remembered workspace + discovery |
| GET  | `/api/battles` | list prepared battles from `Battle_Library.xlsx` (counts + warnings) |
| POST | `/api/battles/<id>/load` | resolve a battle definition into a live encounter (cross-workbook) |
| GET  | `/api/sources` | list creature sources (tab/name/AC/HP) from Combat Tracker for the builder |
| GET  | `/api/battles/<id>/definition` | read a battle's meta + roster rows for editing |
| POST | `/api/battles/save` `{battle}` | create/update a battle definition (creates the library if absent) |
| POST | `/api/battles/<id>/duplicate` · DELETE `/api/battles/<id>` | duplicate / remove a battle |
| GET  | `/api/party-profiles` · GET `/api/party-profiles/<id>` | list / read party profiles (Game_Data.xlsx) |
| POST | `/api/party-profiles/save` · DELETE `/api/party-profiles/<id>` | create-update / remove a party profile |
| POST | `/api/resume` | reload `encounter_state.json` |
| GET  | `/api/state` | full enriched encounter JSON |
| POST | `/api/initiative/roll` `{scope: all\|npcs\|pcs}` | roll initiative |
| POST | `/api/initiative/set` / `/api/initiative/sort` | set one / sort & start |
| POST | `/api/turn/next` · `/api/turn/prev` · `/api/round/next` | turn/round advance |
| POST | `/api/combatant/<id>/damage` · `/heal` · `/temp-hp` | HP events |
| POST | `/api/combatant/<id>/stats` `{ac?, max_hp?}` | manual AC/HP edit (not exported) |
| POST | `/api/combatant/<id>/status` · `/notes` | status override / notes |
| POST | `/api/combatant/<id>/member/<mid>/status` · `/can-attack` | per-member edits |
| POST | `/api/attack/roll` · `/api/attack/apply` | roll an attack / write its damage |
| POST | `/api/group/<id>/allocate` `{mode}` | focused/frontload/even group damage |
| POST | `/api/group/<id>/attack` `{mode: batch\|mob\|average, ac_override?}` | group attack |
| POST | `/api/spell/roll-damage` | roll a dice expression (no mutation) |
| POST | `/api/spell/aoe` | apply AoE damage to many targets (saves, resist, groups) |
| POST | `/api/spell/mass-heal` | heal many targets at once |
| POST | `/api/condition` `{add?\|remove?}` | conditions |
| POST | `/api/limited-use/consume` · `/restore` · `/add` | limited-use counters |
| POST | `/api/log/undo` · DELETE/PATCH `/api/log/<id>` | log edits (drive HP) |
| POST | `/api/export-excel` `{path}` | write current state to a workbook copy |
| GET  | `/api/debug-info` | runtime/template diagnostics |

`@require_encounter` guards routes needing a loaded encounter; `@mutate` runs the handler,
autosaves, and returns the enriched JSON.

## 7. Frontend structure

- **setup.html** — load/resume, roster preview table with inline-editable AC / Max HP /
  initiative, initiative roll buttons, Start Encounter.
- **encounter.html** — the dashboard. Card renderers: `renderPCCard`, `renderMonsterCard`,
  `renderGroupCard`, `renderMemberRow`. Modals: attack, group-attack, condition, stat-block,
  AoE, Mass Heal, dice-roller. Shared helpers: `healTargetsFor`/`healTargetOptions`,
  `dmgTargetOptions` (annotates options with AC), `parseAmount`, `escHtml`/`escJs`,
  `renderDamageResult`. State lives in `ENC` (the last fetched encounter); every mutation
  re-fetches and re-renders.

## 8. Persistence & Excel I/O

- Autosave to `data/encounter_state.json` after every mutation; "Resume" reloads it on startup.
- "Load from Excel" re-imports (warn before discarding live state). Loader contract and string
  formats are documented in [MONSTER_CARD_TEMPLATE.md](MONSTER_CARD_TEMPLATE.md). A Battle List
  row with **Quantity 0** (or negative) is skipped entirely; blank Quantity defaults to 1.
- "Export to Excel" writes current HP/status/log into a **copy** of the workbook; the source is
  never overwritten implicitly. Manual AC/HP edits are state-only and not part of the round-trip.

## 9. Operational notes / gotchas

- **Restart after backend changes.** The server runs with `use_reloader=False` (to avoid a
  port-holding child process), so edits to `app.py` or `engine/*.py` require a manual restart.
  Template/JS/CSS changes show up on a browser refresh (the encounter page is served no-cache).
  Symptom of a stale server: a new UI control appears but its action 404s / does nothing.
- **Port cleanup.** On startup `app.py` kills any stale process holding port 5000.
- **PCs usually have no `max_hp` / `ac`.** There are no PC character sheets; PCs are tracked by
  a damage/heal count-up and DM-entered AC. Code that targets "anything healable/attackable"
  must not filter PCs out by missing `max_hp` — heal targeting and AC-override paths account for
  this.
- **Roster edits need Load, not Resume.** The editable AC/HP fields live on the roster page,
  which only appears after "Load Workbook"; "Resume" jumps straight to the encounter page.

## 10. Non-goals (unchanged from v1)

No automatic PC hit determination, no full character sheets, no spell-targeting/save
auto-resolution, no condition-rule automation, no grid/map or movement math, no reaction
auto-resolution, no multi-user/networked play. The DM makes all rules and narrative calls.
