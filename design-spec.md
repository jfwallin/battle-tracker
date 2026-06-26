SPEC — "Encounter Engine": A DM Combat Tracker (Flask + HTML)
1. Purpose
A local web app that runs a D&D 5e combat encounter from an existing Excel workbook. It replaces a formula-based spreadsheet tracker that can't roll dice on click, can't apply attacks to targets, and requires too much manual turn/round entry. The app must read the existing Excel format as its source of truth for setup, then provide a fast, click-driven combat interface with automatic turn/round advancement, dice rolling, and damage application.

Design priority: speed and low bookkeeping during live play. Assist DM judgment; never auto-adjudicate full 5e rules.

2. Tech Stack
Backend: Python 3.11+, Flask. openpyxl for Excel read/write, pandas optional for table parsing.
Frontend: Server-rendered Jinja2 templates + vanilla JS (or Alpine.js for reactivity — no heavy framework required). WebSocket optional; polling/AJAX is fine for single-user local use.
State: In-memory encounter object, autosaved to a JSON file (encounter_state.json) after every mutation. Excel is read on load and written only on explicit "Export to Excel."
Run mode: flask run, single user, localhost. No auth.
3. Architecture / File Structure
encounter-engine/ app.py # Flask app, routes engine/ excel_loader.py # parse workbook -> Encounter model excel_exporter.py # Encounter model -> workbook models.py # Combatant, Group, Member, Attack, LogEvent, Encounter dice.py # dice parser + roller combat.py # turn/round logic, damage application, attack pool static/ (css, js) templates/ (index.html, partials) data/ sample_encounter.xlsx encounter_state.json tests/
4. Excel Input Schema (the contract the loader must parse)
4a. Battle List sheet — encounter roster
An Excel table named Encounter starting at A1, header row 1:

Col	Field	Notes
A	Type	PC / NPC / Ally / Hazard / Event
B	Name	display name
C	NPC Source	name of the creature sheet tab (blank for PCs/hazards)
D	Quantity	integer; >1 means a group
E	Group Name	optional label, e.g. "Wisp Pack"
F	Initiative Mode	Individual / Shared / Grouped Attacks / Mob
G	HP Override	optional; if set, use instead of sheet Max HP
H	Notes	free text
4b. Creature sheets (e.g. Korrum, Veyrath) — fixed cell map
Every creature tab follows one template. The loader reads by fixed cell address:

Field	Cell
Name	B3
Epithet	B4
Role	B5
Size/Type/Alignment	B6
AC	B7
Max HP	B8
Speed	B9
Proficiency Bonus	B10
Initiative Mod	B11
Ability scores (STR..CHA)	B15:G15
Ability modifiers	B16:G16
Save bonuses	B18:G18
Resistances / Immunities / Cond. Immunities / Senses	B21 / B22 / B23 / B24
Attacks (one per row)	rows 28–29: A=name, B=type, C=to-hit, D=reach/range, E=damage, F=save, G=effect
Traits	A/B rows 33–36
Actions	A=name, B=recharge/uses, C=save, D=description, rows 40–42
Bonus actions / Reactions / Bloodied	rows 46 / 50 / 54
Attacks per round (tracker hook)	K2
Parser robustness: attacks/traits/actions are variable-length — read downward from the section header row until the name column is blank, rather than assuming exactly 2 attack rows. Treat the section header labels (ATTACKS, TRAITS, ACTIONS, BONUS ACTIONS, REACTIONS, BLOODIED) as anchors so the layout can flex.

4c. Damage/to-hit string formats to parse
To-hit: +10, 10, or — (no attack roll).
Damage: "2d8+6 fire", "3d10+7 bludgeoning", possibly multiple clauses "1d8+3 slashing plus 2d6 fire".
Save: "DC 17 STR", "DC 16 DEX", or —. The dice engine must parse these into structured objects.
5. Runtime Data Model (JSON)
Encounter {
  round: int,
  turnIndex: int,            // index into initiative order
  order: [combatantId,...],  // sorted desc by initiative, ties stable
  combatants: {
    id: {
      id, name, type,        // PC|NPC|Ally|Hazard|Event
      sourceTab,             // creature sheet name or null
      initiative: int|null,  // editable
      initiativeMod: int,
      ac: int|null,
      maxHp: int|null, tempHp: int,
      conditions: [string],
      concentration: string|null,
      status: "Active|Bloodied|Down|Dead|Fled|Removed", // derived but overridable
      attacks: [Attack],
      limitedUse: [{name, max, used}],   // legendary res, recharge, slots
      isGroup: bool,
      members: [Member]|null,            // for groups
      initiativeMode: "Individual|Shared|Grouped Attacks|Mob",
      notes: string
    }
  }
}
Member { id, name, maxHp, damage, conditions, canAttack: bool }   // currentHp = maxHp - damage (+heal)
Attack { name, type, toHit:int|null, damageDice:[{n,die,bonus,type}], saveDc:int|null, saveAbility, reach, effect }
LogEvent { id, round, turn, source, target, eventType, amount, damageType, attackName, notes, ts }
currentHp for individuals and members is always derived from maxHp minus the sum of damage events plus healing events in the log (the log is the single source of truth), with a manual-override path that writes an adjustment event.

6. Dice Engine (dice.py)
parse_dice("2d8+6 fire") -> {n:2, die:8, bonus:6, type:"fire"} (and a list when multiple clauses).
roll(n, die) -> [individual rolls]; sum + bonus.
roll_attack(toHit, targetAc, {adv, dis}) -> {d20, nat, total, hit:bool, crit:bool} — nat 20 always hits & crits, nat 1 always misses; otherwise total >= AC. Advantage/disadvantage rolls 2d20.
roll_damage(damageDice, crit:bool) — crit doubles dice count, not the flat bonus.
roll_batch(count, toHit, targetAc) — returns array of attack results + hit count (Batch mode).
mob_hits(attackers, toHit, targetAc, override=None) — expected-hits via this table (required d20 → attackers per hit): 1–5→1, 6–12→2, 13–14→3, 15–16→4, 17–18→5, 19→10, 20→20. requiredRoll = clamp(AC - toHit, 1, 20); hits = floor(attackers / perHit); override wins if provided. Every roll returns the individual dice so the UI can show them and the DM can overrule.
7. Functional Requirements
7.1 Initiative & Turn Tracker (fixes the spreadsheet's biggest pain)
"Roll All NPC Initiative" button → rolls d20 + initiativeMod per NPC (one shared roll for Shared/Grouped/Mob groups). Results are stored as static values (no re-roll on other actions — this was the spreadsheet bug).
PC initiative entered manually (inline-editable field) or "Roll PCs" for NPC-style auto-roll.
Auto-sort descending; stable tie-break (DM can drag to reorder ties).
"Next Turn" / "Previous Turn" buttons: advance turnIndex; when wrapping past the last combatant, auto-increment Round and reset to top. "Next Round" also available. The current combatant is highlighted and scrolled into view. No manual turn-number typing required.
On a combatant's turn start: reset their per-turn flags (reaction used, group attack commitments), decrement/check recharge & condition durations (prompt, don't auto-resolve).
7.2 HP, Damage & Healing
Each combatant card has Apply Damage / Apply Healing inputs + buttons → writes a LogEvent; HP recomputes instantly.
Temp HP handled (damage depletes temp first).
Status auto-derives: >50% Active, ≤50% Bloodied, ≤0 Down/Dead; DM can override to Fled/Removed.
Undo last event and edit/delete any log row → HP recalculates.
7.3 Attacks (the "connect attacks to damage" requirement)
Each combatant shows its parsed attacks as buttons. Clicking "Roll Longsword":
Prompts/【selects】 a target (dropdown of combatants).
Rolls to-hit vs target AC (with adv/dis toggle), shows d20 + total + hit/crit.
On hit, rolls damage (crit-aware), shows dice.
"Apply to target" writes the damage LogEvent.
Save-based attacks: show DC + ability; DM clicks pass/fail per target; apply full/half.
Multiattack/action packages: a creature can define multiple attacks; DM picks which to use this turn.
7.4 Groups (individual / batch / mob)
Group card shows Active/Original count, total HP, attack pool (sum of attacks-per-round over members that are alive AND canAttack). Dead/incapacitated members auto-drop from the pool.
Expandable member list with per-member HP, status, can-attack toggle, conditions.
Damage allocation modes: Focused (pick a member), Front-loaded (auto-cascade to lowest-HP member, overflow to next — this the app CAN do, unlike the spreadsheet), Even split, Casualty-threshold (mob).
Attack modes: Individual (roll each), Batch (roll N d20s, show compact hits/damage), Mob (expected-hits table + override), Average damage (hits × avg).
Attack-pool commitment tracker: assign attacks to targets during the group's turn; "Remaining" counts down; resets at the group's next turn.
7.5 Conditions, Limited-Use, Hazards
Conditions: multi-select chips per combatant (Restrained, Stunned, Prone, Frightened, Grappled, Blinded, Invisible, Poisoned, Concentrating, Incapacitated, Unconscious, Dead, Fled) + free-text duration note. No auto-rules enforcement.
Limited-use abilities: show counters only for creatures that have them (legendary resistance, recharge X–Y, spell slots, reaction-used). Click to consume/restore.
Hazards/Events: non-creature rows with initiative + notes + optional damage/save, no HP.
7.6 Combat Log
Append-only, newest-first, filterable by round/combatant. Each row: round, turn, source, target, type, amount, damage type, attack/effect, notes. Drives all HP. Editable/undoable.
7.7 Persistence
Autosave full state to encounter_state.json after every mutation; reload on startup ("Resume encounter").
"Load from Excel" (re-import roster + stat blocks; warn before discarding live state).
"Export to Excel" writes current HP/status/log back into a copy of the workbook (don't overwrite the source unless confirmed).
8. REST API (Flask routes)
GET / -> main dashboard POST /api/load-excel {path} -> import encounter POST /api/initiative/roll {scope: all|npcs|pcs} POST /api/initiative/set {combatantId, value} POST /api/turn/next | /turn/prev | /round/next POST /api/combatant/:id/damage {amount, type, source, attackName} POST /api/combatant/:id/heal {amount} POST /api/attack/roll {combatantId, attackName, targetId, adv, dis} POST /api/attack/apply {result} -> writes log event POST /api/group/:id/allocate {mode, amount, memberId?} POST /api/group/:id/attack {mode, ...} POST /api/condition {combatantId, add?, remove?} POST /api/log/undo | /log/:id (PATCH|DELETE) POST /api/export-excel {path} GET /api/state -> full encounter JSON
All mutating routes return the updated encounter JSON so the UI re-renders.

9. UI Layout
Top banner: Round #, Current Turn (big, highlighted), Next Up, encounter summary (enemies alive, total enemy HP, attacks in pool). [Next Turn] is the primary, always-visible button.
Center: initiative-ordered combatant cards; current card highlighted. Each card: name, init, AC, HP bar (current/max + temp), status pill, condition chips, attack buttons, quick damage/heal input. Group cards expand to member rows.
Right rail / modal: dice results feed (last rolls with individual dice), combat log.
Setup screen: load Excel, review roster, choose per-group initiative/attack modes, roll initiative, Start Encounter.
Keyboard shortcuts: N next turn, D damage current target, R roll selected attack.
10. Non-Goals (v1)
No automatic PC hit determination, no full character sheets, no spell-targeting/save auto-resolution, no condition-rule automation, no grid/map, no movement/range math, no reaction auto-resolution, no multi-user/networked play. The DM makes all rules/narrative calls.

11. Build Milestones
M1 — Excel loader + models: parse Encounter table + creature sheets into the data model; unit tests on sample_encounter.xlsx.
M2 — Dice engine: parser + roller + tests (to-hit, crit, batch, mob table).
M3 — Turn/round core + dashboard: initiative roll/sort, Next/Prev turn with auto-round, current-turn highlight, autosave JSON.
M4 — HP + combat log: damage/heal via log, undo, status derivation.
M5 — Attacks: attack buttons → to-hit → damage → apply to target.
M6 — Groups: member tracking, attack pool, allocation modes, batch/mob.
M7 — Conditions, limited-use, hazards.
M8 — Excel export + polish + shortcuts.
12. Acceptance Criteria
During a live encounter the DM can, each in ≤2 clicks: see whose turn it is and who's next; advance the turn with the round auto-incrementing; roll any creature's attack against a chosen target and apply the damage; see remaining enemies, their HP, and a group's remaining attack pool (with dead members excluded); add/clear a condition; undo the last action; and export the final state to Excel. Initiative, once rolled, never changes on unrelated actions.