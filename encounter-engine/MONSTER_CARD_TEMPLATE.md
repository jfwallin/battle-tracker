# Monster Card Authoring Guide

How to add new creatures to an Encounter Engine workbook so the loader
([`engine/excel_loader.py`](engine/excel_loader.py)) parses them correctly. Written for a
human or an AI assistant generating stat blocks. Follow the cell addresses and string
formats **exactly** — the loader reads by fixed cell, not by label matching.

A workbook has two kinds of sheets:

1. **`Battle List`** — one roster row per combatant (who is in the fight).
2. **Creature tabs** — one sheet per creature type, holding the stat block. Referenced by
   name from the Battle List's `NPC Source` column.

---

## 1. The `Battle List` sheet (roster)

A sheet **named exactly `Battle List`**. The loader finds the header row by scanning column A
for a cell whose value is `Type` (case-insensitive); rows below it are read until a row has a
blank column A. Columns:

| Col | Field            | Notes |
|-----|------------------|-------|
| A   | Type             | `PC` / `NPC` / `Ally` / `Hazard` / `Event`. **Required** — a blank cell ends/skip the row. |
| B   | Name             | Display name. For a group this is the per-member base name (e.g. `Flame Wisp` → `Flame Wisp 1`, `Flame Wisp 2`, …). |
| C   | NPC Source       | Name of the creature tab to pull stats from. Blank for PCs/Hazards/Events. Must match a sheet/tab name exactly. |
| D   | Quantity         | Integer. `>1` makes this a **group** with that many members. Blank/invalid = 1. |
| E   | Group Name       | Optional label shown instead of Name (e.g. `Wisp Pack`). |
| F   | Initiative Mode  | `Individual` / `Shared` / `Grouped Attacks` / `Mob`. Blank = `Individual`. `Shared`/`Grouped Attacks`/`Mob` roll one shared initiative for the whole group. |
| G   | HP Override      | Optional integer. If set, overrides the creature tab's Max HP. |
| H   | Notes            | Free text. |

**Notes**
- A PC row needs no creature tab; leave column C blank and set AC/HP via play or HP Override.
- `Hazard`/`Event` rows are non-creatures (initiative + notes, no HP) — leave C blank.

---

## 2. Creature tabs (stat block)

One sheet per creature type. The **tab name must match** the `NPC Source` value in the Battle
List. The loader reads fixed cells plus a few **section anchors**.

### 2a. Fixed cells

| Field                      | Cell    | Notes |
|----------------------------|---------|-------|
| Name                       | `B3`    | |
| Epithet                    | `B4`    | |
| Role                       | `B5`    | |
| Size / Type / Alignment    | `B6`    | e.g. `Large elemental, neutral` |
| AC                         | `B7`    | integer |
| Max HP                     | `B8`    | integer (overridden by Battle List col G if set) |
| Speed                      | `B9`    | text, e.g. `30 ft., fly 60 ft.` |
| Proficiency Bonus          | `B10`   | integer |
| Initiative Mod             | `B11`   | integer (added to a d20 on "Roll Initiative") |
| **Attacks per round**      | `K2`    | integer; drives a group's attack pool. Defaults to 1 if blank. |

### 2b. Ability rows (columns B–G = STR, DEX, CON, INT, WIS, CHA)

| Row  | Field            |
|------|------------------|
| 15   | Ability scores   |
| 16   | Ability modifiers|
| 18   | Save bonuses     |

So `B15:G15` are the six scores in STR, DEX, CON, INT, WIS, CHA order, `B16:G16` the mods,
`B18:G18` the save bonuses. Blank cells read as 0.

### 2c. Defenses / senses

| Field                 | Cell  |
|-----------------------|-------|
| Resistances           | `B21` |
| Immunities            | `B22` |
| Condition Immunities  | `B23` |
| Senses                | `B24` |

Resistances/Immunities are matched as **substrings** against an incoming damage type (e.g. an
AoE of type `fire` is resisted if `B21` contains the word `fire`). Use comma-separated lowercase
damage words for reliable matching, e.g. `fire, poison`.

### 2d. Section anchors (variable-length lists)

These sections are found by scanning column A for a header cell whose text (uppercased, trimmed)
matches the anchor. Rows are then read **downward from the row below the header until column A is
blank or the next anchor header is reached** — so a creature can have any number of rows.

Anchors: `ATTACKS`, `TRAITS`, `ACTIONS`, `BONUS ACTIONS`, `REACTIONS`, `BLOODIED`.

**ATTACKS** — one attack per row, columns A–G:

| Col | Field        | Format |
|-----|--------------|--------|
| A   | Name         | Attack name. (A row literally named `Name` is treated as a template header and hidden in the UI.) |
| B   | Type         | e.g. `Melee Weapon`, `Ranged Spell` |
| C   | To-hit       | `+10`, `10`, or `—`/blank for no attack roll (save-only) |
| D   | Reach/Range  | e.g. `5 ft.`, `60 ft.` |
| E   | Damage       | see **String formats** below |
| F   | Save         | `DC 17 STR`, `DC 16 DEX`, or `—`/blank |
| G   | Effect       | free text rider (e.g. `Prone on fail`) |

**TRAITS / ACTIONS / BONUS ACTIONS / REACTIONS / BLOODIED** — read as name/description pairs:
column A = name, column B = description.

---

## 3. String formats the parser requires

- **To-hit** (`C` on attack rows): `+10`, `10`, or `—`/blank (= no attack roll).
- **Damage** (`E`): `<n>d<die>[+/-bonus] <type>`, e.g. `2d8+6 fire`. Chain multiple clauses with
  the literal word **`plus`**: `1d8+3 slashing plus 2d6 fire`. `—`/blank = no damage.
- **Save** (`F`): `DC <number> <ABILITY>` where ability is a 3-letter code
  (`STR`/`DEX`/`CON`/`INT`/`WIS`/`CHA`), e.g. `DC 17 STR`. `—`/blank = no save.
- Use the dash `—` (or blank) to mean "none" in any to-hit/damage/save cell.

---

## 4. Rules for an AI generating a creature tab

1. Put the tab name = the exact `NPC Source` you'll use in the Battle List.
2. Fill the fixed cells above; never shift them. AC (`B7`) and Max HP (`B8`) must be integers.
3. Lay out abilities across columns B–G in STR, DEX, CON, INT, WIS, CHA order on rows 15/16/18.
4. Write the literal uppercase anchor word (`ATTACKS`, `TRAITS`, …) in column A above each
   section; put real rows immediately below with **no blank rows inside a section** (a blank
   column-A cell ends the section).
5. Use the exact string formats in §3. Lowercase damage/condition words help substring matching.
6. Set `K2` to the number of attacks the creature makes per round (used for group attack pools).
7. Add a matching `Battle List` row (Type, Name, NPC Source = tab name, Quantity, Init Mode).

---

## 5. Worked example

**Battle List row**

| A    | B           | C            | D | E         | F      | G | H            |
|------|-------------|--------------|---|-----------|--------|---|--------------|
| NPC  | Flame Wisp  | Flame Wisp   | 3 | Wisp Pack | Mob    |   | Swarm caster |

**Creature tab named `Flame Wisp`**

```
        A              B          C       D        E              F            G
3   Name           Flame Wisp
4   Epithet        Ember Mote
5   Role           Skirmisher
6   Size/Type      Small elemental, chaotic neutral
7   AC             13
8   Max HP         22
9   Speed          0 ft., fly 40 ft. (hover)
10  Prof Bonus     2
11  Init Mod       3
                   STR        DEX     CON     INT     WIS     CHA
15  Scores         6          16      12      6       10      11
16  Mods           -2         3       1       -2      0       0
18  Saves          -2         5       1       -2      0       0
21  Resistances    fire
22  Immunities     poison
23  Cond Imm.      poisoned
24  Senses         darkvision 60 ft.
...
28  ATTACKS
29  Ember Touch    Melee Spell  +5      5 ft.   1d6+3 fire     —            Ignites on crit
30  Spark Burst    Ranged Spell —       20 ft.  2d6 fire       DC 12 DEX    Half on save
...
33  TRAITS
34  Ignition       When it dies, deals 1d6 fire to adjacent creatures.
...
40  ACTIONS
41  Flare          Recharge 5-6 — — DC 12 DEX 3d6 fire in 10 ft.
```

(`K2` = `2`, since the wisp makes two attacks per round.)
