# Encounter Engine Expansion — Intent and Staged Design Direction

## Purpose

Extend the encounter engine from a single-workbook combat tracker into a reusable encounter-management tool.

The goal is not to turn the app into a full virtual tabletop or complete rules engine. The goal is to make encounter preparation, loading, modification, live combat, and post-combat review substantially easier while preserving the current strengths:

* Fast live use
* Low bookkeeping burden
* Clear initiative and turn flow
* Flexible DM judgment
* Reliable log-derived HP
* Support for individual creatures, groups, mobs, hazards, and events

The app must remain generic. It must work for one-shots, short adventures, long-running games, published material, and entirely custom content.

## Critical Compatibility Requirement

**Preserve the existing `Combat Tracker.xlsx` workbook.**

`Combat Tracker.xlsx` is the current monster/NPC library and the established Excel input contract for the app. Do not redesign its structure, replace its stat-block tabs, migrate its monster data into a new workbook, or require existing creature sheets to be reformatted.

All new functionality must treat `Combat Tracker.xlsx` as the authoritative source for existing monster and NPC definitions.

The app may add companion workbooks such as `Battle_Library.xlsx` and optional `Game_Data.xlsx`, but those files must reference creatures already defined in `Combat Tracker.xlsx`.

Changes to `Combat Tracker.xlsx` should be limited to backward-compatible additions only when strictly necessary. Existing monster tabs, current battle-list data, and current encounter loading behavior must continue to work without modification.

A workspace containing only `Combat Tracker.xlsx` must continue to support the current single-workbook workflow.

---

# Guiding Intent

## 1. Separate reusable definitions from live combat

A live encounter should not be the only place where creatures, party members, battle setups, and modifications exist.

The system needs three conceptual layers:

1. **Reference data**
   Reusable monster and NPC definitions from the existing `Combat Tracker.xlsx` workbook.

2. **Battle definitions**
   Reusable encounter setups describing who and what should be present.

3. **Live encounter state**
   The current initiative, HP, conditions, actions, damage log, and all other facts created during play.

The user should be able to load the same battle definition repeatedly without carrying over damage, casualties, initiative, or conditions from a prior run.

At the same time, the user should be able to save a live battle as a new encounter when continuity is desired.

---

## 2. Prefer a standard data directory over repeated file selection

The normal workflow should use a selected or configured directory containing Excel files with predictable names.

This is important for both testing and normal use. It reduces repeated file-picker steps and makes it easier to move or back up an entire game data set.

Example directory:

```text
Encounter_Data/
    Combat Tracker.xlsx
    Battle_Library.xlsx
    Game_Data.xlsx              # optional
    Completed_Battles/
    Encounter_Snapshots/
```

The app should allow the user to select a workspace directory once and remember it as the current workspace.

The app should then automatically look for standard filenames:

```text
Combat Tracker.xlsx
Battle_Library.xlsx
Game_Data.xlsx
```

The app must clearly report missing optional files and missing required files.

Example:

> Combat Tracker found.
> Battle Library found.
> Game Data not found; continuing without party profiles, recurring NPCs, or reusable variants.

The app should still support a manual file-selection mode for testing, migration, or nonstandard setups.

---

# Data Workbooks

## 1. Existing Combat Tracker Workbook

### Intent

`Combat Tracker.xlsx` remains the current source workbook for monster and NPC stat blocks.

It is not being redesigned in this project.

The new battle-library and optional game-data features must work with its current tabs, current card format, existing battle-list data, and current loader behavior.

A source creature or NPC may be identified through the existing supported source-tab mechanism. Where practical, the app may add a stable source identifier as a backward-compatible enhancement, but it must not require existing stat-block sheets to be rewritten before they can load.

### Example use

A battle definition may reference a creature or NPC source:

```text
Hobgoblin Soldier
```

The live encounter then creates individual combatants or group members such as:

```text
Hobgoblin Soldier 1
Hobgoblin Soldier 2
Hobgoblin Soldier 3
```

Those active combatants may later have different HP, conditions, status, attacks available, and notes. None of that changes the original creature source unless the user explicitly edits the source workbook.

### Existing data responsibilities

`Combat Tracker.xlsx` remains responsible for reusable creature and NPC information, including:

| Information          | Intent                                                       |
| -------------------- | ------------------------------------------------------------ |
| Creature or NPC name | Human-readable source name                                   |
| AC                   | Default armor class                                          |
| Max HP               | Default maximum HP                                           |
| Initiative Modifier  | Default initiative bonus                                     |
| Ability Scores       | Strength through Charisma, where available                   |
| Save Bonuses         | Specific saving throws, where available                      |
| Attacks              | Attack name, to-hit, damage, range, effects                  |
| Traits and Actions   | Readable stat-block information                              |
| Limited Uses         | Recharge abilities, spell slots, legendary resistances, etc. |
| Notes                | Optional tactical or rules reminders                         |

Optional fields should remain optional. The app should continue to support simple creatures with only AC, HP, initiative, and one attack.

### Important behavior

The existing workbook should remain editable through direct Excel editing exactly as it is now.

Future app-based library editing may be added later, but it must preserve the existing source format and not force a migration to a new workbook layout.

---

## 2. Battle Library Workbook

### Intent

`Battle_Library.xlsx` stores reusable battle definitions.

A battle definition is a recipe for creating an encounter. It identifies participants, groups, hazards, events, starting conditions, and setup notes. It does not contain the actual results of a played battle.

A user should be able to prepare several battles before a session, then choose one from a list and load it quickly.

### Required workbook contents

The Battle Library should contain:

1. A battle index.
2. One battle-definition sheet or structured table per saved battle.
3. Optional reusable battle components, such as common events or terrain effects.

### Battle index intent

The battle index is the app’s entry point for prepared encounters.

It should allow the app to display a selectable list of battles with enough context to choose the correct one quickly.

Minimum information:

| Field                 | Intent                                           |
| --------------------- | ------------------------------------------------ |
| Battle ID             | Stable unique identifier                         |
| Battle Name           | Human-readable name                              |
| Definition Location   | Sheet or table containing the battle setup       |
| Status                | Draft, Ready, Test, Archived, Completed Template |
| Tags                  | Optional filtering and searching                 |
| Notes                 | Brief DM-facing description                      |
| Default Party Profile | Optional party profile to import                 |
| Last Modified         | Optional management information                  |

### Battle definition intent

A battle definition should describe what the encounter starts with.

Each entry should identify one participant, group, hazard, event, reinforcement, or other encounter element.

Minimum information:

| Field                     | Intent                                                    |
| ------------------------- | --------------------------------------------------------- |
| Entry Type                | PC, NPC, Ally, Hazard, Event, Reinforcement               |
| Source Name or Source Tab | Reference to a creature or NPC in `Combat Tracker.xlsx`   |
| Display Name              | Name shown in the live encounter                          |
| Quantity                  | Number of identical entries                               |
| Group Name                | Optional group identity                                   |
| Initiative Mode           | Individual, Shared, Grouped Attacks, Mob                  |
| Variant ID                | Optional reusable modification                            |
| Starting HP               | Optional override                                         |
| Starting Conditions       | Optional conditions at encounter start                    |
| Starting Status           | Optional status such as Active, Hidden, Fled, Unconscious |
| Notes                     | Tactical or setup note                                    |
| Trigger                   | Optional delayed-entry or activation condition            |

### Important behavior

Loading a battle must resolve the battle definition into a self-contained live encounter.

The loaded encounter should preserve the exact resolved values used at the moment it was loaded. Later edits to `Combat Tracker.xlsx`, `Battle_Library.xlsx`, or `Game_Data.xlsx` must not silently alter a battle already in progress.

---

## 3. Game Data Workbook

### Intent

`Game_Data.xlsx` is optional.

It exists only for reusable material that is not a generic monster template and not limited to a single prepared battle.

The app must work without this file.

The Game Data workbook is most useful for:

* Party profiles
* Recurring named NPCs
* Reusable monster or NPC variants
* Optional reusable custom content

It should not become a miscellaneous storage location for active combat state.

### Recommended contents

## Party Profiles

### Intent

A party profile is a reusable roster that can be imported into prepared or dynamically created battles.

It is not a full player-character management system.

It exists to avoid repeatedly typing names and basic combat reference values.

Minimum information:

| Field               | Intent                              |
| ------------------- | ----------------------------------- |
| Profile ID          | Stable identifier for the roster    |
| Character Name      | Name shown in combat                |
| Type                | PC, ally, companion, familiar, etc. |
| Default AC          | Optional starting AC                |
| Default Max HP      | Optional reference HP               |
| Initiative Modifier | Optional initiative bonus           |
| Notes               | Optional reminders                  |

A battle definition may specify a default party profile, but the DM must be able to change, omit, or supplement it during battle loading.

## Recurring NPCs

### Intent

Recurring NPCs represent named individuals who may appear across multiple battles.

They are distinct from generic templates because they have stable identity and may have reusable modifications or notes.

Minimum information:

| Field                          | Intent                                           |
| ------------------------------ | ------------------------------------------------ |
| NPC ID                         | Stable unique identifier                         |
| Base Source Name or Source Tab | Underlying `Combat Tracker.xlsx` creature source |
| Display Name                   | Named individual                                 |
| Variant ID                     | Optional reusable modification                   |
| Notes                          | Persistent DM notes                              |
| Status                         | Optional persistent game-level status            |

The app should not automatically propagate battle injuries, deaths, or resource expenditure into recurring NPC records. That should require an explicit user decision after combat.

## Variants

### Intent

Variants are reusable modifications layered onto a base creature or NPC template.

They allow the DM to create meaningful variations without duplicating whole stat blocks.

Examples include:

* Armored version
* Elite version
* Wounded version
* Fire-themed caster
* Veteran archer
* Guard captain
* Creature carrying a special item

A variant may change AC, HP, attacks, abilities, resistances, limited uses, notes, or other supported stat-block fields.

Variants should be reusable across multiple battles.

### Important behavior

Variants should not alter the original template.

The app should resolve data in this order:

1. Base creature template from `Combat Tracker.xlsx`
2. Reusable variant from `Game_Data.xlsx`
3. Battle-specific override from `Battle_Library.xlsx`
4. Individual encounter override
5. Live combat changes

## Custom Content

### Intent

Custom Content is optional reusable material that is not naturally represented as a creature.

Examples:

* Hazards
* Lair actions
* Events
* Timed magical effects
* Terrain effects
* Reinforcement packages
* Environmental turns

This should be included only if it proves useful in repeated encounters. A battle should also be able to define a local hazard or event directly without requiring it to become reusable global content.

---

# Live Encounter State

## Intent

The live encounter remains the place where combat actually happens.

It includes:

* Initiative values
* Turn order
* Current round
* Current HP
* Temp HP
* Conditions
* Status
* Attack availability
* Reaction/action tracking
* Group member state
* Damage and healing events
* Notes created during play

This remains distinct from all Excel source data.

The current log-derived HP model should remain the source of truth for live HP and should continue to support undo, edit, and correction of prior events.

---

# App Changes Needed

## 1. Workspace Selection and Validation

### Intent

The app should understand a workspace or data directory instead of assuming one manually selected workbook.

### Required user experience

At startup, the user should be able to:

* Select a workspace directory.
* See which expected files were found.
* Choose a different directory.
* Continue with optional files missing.
* Open a diagnostic summary if loading fails.

The app should validate:

* `Combat Tracker.xlsx` availability
* `Battle_Library.xlsx` availability, when prepared battles are desired
* Optional `Game_Data.xlsx` availability
* Required sheets or structured tables in companion workbooks
* Broken references between workbooks
* Invalid quantities
* Invalid group or initiative modes

The app should provide helpful, specific errors rather than generic load-failure messages.

The app must continue to allow a user to load `Combat Tracker.xlsx` directly and use the existing workflow even if no workspace directory or companion workbook is present.

---

## 2. Battle Selection

### Intent

The user should be able to select a prepared battle quickly from the Battle Library.

### Required user experience

The setup screen should show:

* Battle name
* Tags
* Status
* Notes
* Default party profile
* Number of NPCs, groups, hazards, and events
* Optional preview of participants

The user should be able to:

* Load battle
* Edit battle definition
* Duplicate battle definition
* Create a new battle
* Archive a battle
* Load a saved battle snapshot
* Continue a previously active encounter

---

## 3. Dynamic Battle Builder

### Intent

The user should be able to construct a battle without editing Excel manually.

The result should be saveable as a reusable battle definition in `Battle_Library.xlsx`.

### Required user experience

The battle builder should allow the user to:

* Create a new battle name.
* Select or omit a party profile.
* Search sources from `Combat Tracker.xlsx`.
* Add creatures or NPCs.
* Set quantity.
* Create or assign groups.
* Choose initiative mode.
* Apply variants.
* Set starting HP or status overrides.
* Add hazards and events.
* Add local notes.
* Add delayed reinforcements or simple triggers.
* Save as a new battle definition.
* Duplicate an existing battle and modify it.

The builder should present a readable encounter summary before saving.

---

## 4. Save and Clone Behaviors

### Intent

The app needs clear distinctions between copying a setup, preserving an active encounter, and saving a future continuation.

### Required commands

| Command                          | Intent                                                                                              |
| -------------------------------- | --------------------------------------------------------------------------------------------------- |
| Duplicate Battle Definition      | Copy a prepared encounter setup; reset all live state.                                              |
| Save Live Encounter Snapshot     | Preserve the exact current battle for later resume.                                                 |
| Save Current State as New Battle | Convert current surviving combatants, HP, conditions, and modifications into a new prepared battle. |
| Export Completed Battle          | Create a historical record and player-facing recap without changing source definitions.             |

The interface must make these choices explicit so the DM does not accidentally preserve or erase live state.

---

## 5. Stat-Block Editing and Modifications

### Intent

The DM should be able to make a creature stronger, weaker, altered, or specialized without duplicating a full creature sheet.

### Required user experience

From a stat-card popup or preparation interface, the DM should be able to:

* Roll an ability check.
* Roll a saving throw.
* Roll an attack.
* Roll damage only.
* Roll recharge.
* Edit temporary encounter values.
* Create a local battle-only modification.
* Save a reusable variant intentionally.

Temporary changes made during a live battle should remain encounter-specific unless explicitly saved as a variant or recurring-NPC update.

---

## 6. Combat Event Logging and Battle Report Export

### Intent

The app should record enough structured information to produce a readable recap of combat, including actions that caused no damage.

The goal is a shareable battle scorecard rather than only a damage ledger.

### Required changes to event logging

The event log should be expanded to capture:

* Round
* Initiative position
* Acting combatant
* Action type
* Action name
* Target or targets
* Attack roll or save result
* Hit, miss, success, failure, or partial result
* Damage or healing applied
* Damage type
* Conditions applied or removed
* Use of bonus actions
* Use of reactions
* Use of legendary actions or equivalent actions
* Important notes

The live combat interface should make logging normal actions easy, but it should not force the DM to document every trivial movement or conversation.

### Required export outputs

The app should export two complementary views.

#### Battle Scorecard

A readable sheet with:

* Initiative order vertically
* Rounds horizontally
* Concise narrative summaries in each cell

Example intent:

> Round 2, Mage: casts Fireball; three targets fail Dexterity saves and take 22 fire damage; one target succeeds and takes 11.

> Round 2, Guard Unit: six attacks; two hit the fighter for 17 slashing damage.

> Round 2, Rogue: reaction used to reduce damage by 9.

#### Detailed Event Log

A filterable, chronological sheet containing all recorded events.

The scorecard is for reading and sharing. The detailed log is for review and correction.

### Player-safe export

The app should offer a player-safe export option that excludes:

* Hidden creature HP
* Secret notes
* Unrevealed abilities
* Internal tactics
* Stat-block information not revealed during play

---

# Staged Implementation Plan

## Stage 1 — Workspace and Battle Library Foundation

### Intent

Move from one imported workbook to a workspace containing standard Excel files while preserving the current `Combat Tracker.xlsx` workflow unchanged.

### Deliverables

* Workspace directory selection
* Standard filename discovery
* Load the existing `Combat Tracker.xlsx` unchanged
* Load `Battle_Library.xlsx`
* Battle Index display
* Prepared battle selection
* Resolve battle entries against creature/NPC sources in `Combat Tracker.xlsx`
* Load battle into the existing live encounter engine
* Validation of source references and basic encounter entries
* Clear missing-file and broken-reference messages

### Compatibility requirements

* A workspace containing only `Combat Tracker.xlsx` must continue to support the existing workflow.
* Existing monster tabs must load without modification.
* Existing battle-list data must continue to work.
* `Battle_Library.xlsx` adds prepared encounter selection; it does not replace or restructure `Combat Tracker.xlsx`.

### Out of scope

* Game Data workbook
* Dynamic battle builder
* Persistent variants
* Detailed player-facing report export
* In-app Excel editing

### Success condition

A user can place `Combat Tracker.xlsx` and `Battle_Library.xlsx` in one directory, select that directory once, choose a prepared battle, and start combat without separately loading multiple files.

---

## Stage 2 — Dynamic Battle Builder and Battle Saving

### Intent

Allow users to create, duplicate, and edit battle definitions from inside the app.

### Deliverables

* Create new battle
* Search and add sources from `Combat Tracker.xlsx`
* Set quantities and groups
* Select individual/shared/grouped/mob initiative mode
* Add local hazards and events
* Add encounter notes
* Duplicate a battle definition
* Save new battle definitions to `Battle_Library.xlsx`
* Basic editing of prepared battle definitions
* Save current encounter as a new battle definition

### Out of scope

* Persistent party profiles
* Recurring named NPCs
* Reusable variants
* Advanced trigger systems
* Full battle recap export

### Success condition

A user can create a battle such as “three mages, five guards as one group, and a hazard at initiative 20,” save it, reload it later, and run it without manual Excel editing.

---

## Stage 3 — Game Data and Reusable Variants

### Intent

Add optional reusable party rosters, named NPCs, and stat-block variations.

### Deliverables

* Optional `Game_Data.xlsx`
* Party profile import
* Recurring NPC records
* Reusable variant records
* Apply variants during battle construction
* Create a reusable variant from a stat-card edit
* Explicit save behavior for recurring NPC updates
* Clear resolution order for base source, variants, overrides, and live changes

### Out of scope

* Automatic persistence of injuries, death, or resource expenditure
* Full player character sheets
* Rules automation

### Success condition

A user can select a party profile, add a named NPC based on an existing creature source, apply a reusable elite variant to a group, and preserve the original `Combat Tracker.xlsx` entries unchanged.

---

## Stage 4 — Rich Combat Event Log and Battle Reports

### Intent

Turn the current combat log into a readable and shareable battle record.

### Deliverables

* Expanded action/event logging
* Action, bonus action, and reaction indicators
* Capture hits, misses, saves, and conditions
* Scorecard export by initiative and round
* Detailed chronological event-log export
* Player-safe export mode
* Completed-battle archive output

### Out of scope

* Automatically interpreting every action
* Full narrative generation
* Complete tactical replay
* Automatic rules adjudication

### Success condition

After combat, the DM can export a clear summary showing who acted each round, what attacks or spells were used, which attacks missed, how damage was applied, and when major conditions or reactions mattered.

---

## Stage 5 — Optional Advanced Encounter Tools

### Intent

Support more complex encounters without making them mandatory.

### Possible deliverables

* Reinforcement waves
* Triggered events
* Phase transitions
* Morale reminders
* Retreat and surrender behaviors
* Environmental initiative turns
* Reusable event packages
* Encounter balance summaries
* Persistent NPC updates after battle
* In-app library manager

### Success condition

Complex multi-stage encounters become easier to run, but simple battles remain just as fast and uncluttered as they are now.

---

# Constraints

The system should continue to avoid becoming:

* A full virtual tabletop
* A grid/map manager
* A complete player-character sheet manager
* A full D&D rules engine
* A mandatory campaign-management database
* A system that silently overwrites Excel source data
* A system that requires redesigning `Combat Tracker.xlsx`

The app should remain a local, single-user encounter-management tool whose primary job is reducing DM bookkeeping during preparation and combat.
