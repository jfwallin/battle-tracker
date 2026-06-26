"""Debug: print Slam attack fields as parsed from Excel."""
import sys
sys.path.insert(0, '.')
from engine.excel_loader import load_encounter

enc = load_encounter('../combat tracker.xlsx')
korrum = next(c for c in enc.combatants.values() if c.name == 'Korrum')
for a in korrum.attacks:
    print(f"name={repr(a.name)} to_hit={repr(a.to_hit)} save_dc={repr(a.save_dc)} save_ability={repr(a.save_ability)} damage_dice={a.damage_dice} effect={repr(a.effect[:60] if a.effect else '')}")
