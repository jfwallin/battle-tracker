import sys
sys.path.insert(0, '.')
from engine.excel_loader import load_encounter
enc = load_encounter('../combat tracker.xlsx')
for c in enc.combatants.values():
    if c.combatant_type == 'NPC' or c.is_group:
        print(f'\n{c.name} (type={c.combatant_type}, is_group={c.is_group}):')
        for a in c.attacks:
            print(f'  attack: name={repr(a.name)}, to_hit={a.to_hit}, dice={[(d.n, d.die, d.bonus, d.damage_type) for d in a.damage_dice]}')
