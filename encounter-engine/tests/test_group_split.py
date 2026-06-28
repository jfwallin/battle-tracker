"""Tests for splitting a group's attack pool across multiple targets."""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import app as A
from engine.models import Encounter, Combatant, Member, Attack, DamageDie, make_id


def _setup(pool_members=8):
    enc = Encounter(source_path="x")
    atk = Attack(name="Longbow", attack_type="Ranged", to_hit=5, reach="150 ft.",
                 damage_dice=[DamageDie(n=1, die=8, bonus=2, damage_type="piercing")])
    grp = Combatant(id="g", name="Archers", combatant_type="NPC", source_tab=None,
                    initiative=10, initiative_mod=0, ac=14, max_hp=11, is_group=True,
                    members=[Member(id=make_id(), name=f"A{i}", max_hp=11) for i in range(pool_members)],
                    attacks=[atk], attacks_per_round=1)
    t1 = Combatant(id="t1", name="Rex", combatant_type="PC", source_tab=None,
                   initiative=12, initiative_mod=0, ac=15, max_hp=40)
    t2 = Combatant(id="t2", name="Prag", combatant_type="PC", source_tab=None,
                   initiative=11, initiative_mod=0, ac=13, max_hp=30)
    enc.combatants = {"g": grp, "t1": t1, "t2": t2}
    enc.order = ["g", "t1", "t2"]
    A._encounter = enc
    return A.app.test_client()


def test_split_partitions_pool_with_unique_rolls():
    c = _setup(8)
    r = c.post('/api/group/g/attack', json={
        'attack_name': 'Longbow', 'mode': 'batch',
        'targets': [{'target_id': 't1', 'count': 3, 'ac_override': 15},
                    {'target_id': 't2', 'count': 5, 'ac_override': 13}]})
    res = r.get_json()['result']
    assert res['pool'] == 8 and res['assigned'] == 8
    segs = res['segments']
    assert len(segs) == 2
    s1, s2 = segs
    assert s1['target_id'] == 't1' and s1['count'] == 3
    assert s2['target_id'] == 't2' and s2['count'] == 5
    # Unique per-attacker rolls: one d20 each, with its own pre-rolled damage.
    assert len(s1['batch']['attacks']) == 3 and len(s2['batch']['attacks']) == 5
    assert all('d20' in a and 'damage' in a for a in s1['batch']['attacks'])


def test_over_pool_rejected():
    c = _setup(8)
    r = c.post('/api/group/g/attack', json={
        'attack_name': 'Longbow', 'mode': 'batch',
        'targets': [{'target_id': 't1', 'count': 5, 'ac_override': 15},
                    {'target_id': 't2', 'count': 5, 'ac_override': 13}]})
    assert r.status_code == 400 and 'pool is only 8' in r.get_json()['error']


def test_under_pool_allowed():
    c = _setup(8)
    r = c.post('/api/group/g/attack', json={
        'attack_name': 'Longbow', 'mode': 'batch',
        'targets': [{'target_id': 't1', 'count': 3, 'ac_override': 15}]})
    res = r.get_json()['result']
    assert res['assigned'] == 3 and res['pool'] == 8 and len(res['segments']) == 1


def test_legacy_single_target_shape():
    c = _setup(6)
    r = c.post('/api/group/g/attack', json={
        'attack_name': 'Longbow', 'mode': 'batch', 'target_id': 't1', 'ac_override': 15})
    res = r.get_json()['result']
    # New segments[] present AND legacy top-level keys mirrored.
    assert len(res['segments']) == 1 and res['segments'][0]['count'] == 6
    assert 'batch' in res and res['to_hit'] == 5 and res['target_ac'] == 15


def test_mob_segment():
    c = _setup(12)
    r = c.post('/api/group/g/attack', json={
        'attack_name': 'Longbow', 'mode': 'mob',
        'targets': [{'target_id': 't1', 'count': 6, 'ac_override': 15},
                    {'target_id': 't2', 'count': 6, 'ac_override': 13}]})
    segs = r.get_json()['result']['segments']
    assert len(segs) == 2 and all('mob' in s for s in segs)
    assert segs[0]['mob']['attackers'] == 6
