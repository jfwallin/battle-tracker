"""Tests for the dice engine."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest
from engine.dice import roll, roll_attack, roll_damage, roll_batch, mob_hits, avg_damage
from engine.models import DamageDie


# ── roll ──────────────────────────────────────────────────────────────────────

def test_roll_range():
    for _ in range(200):
        r = roll(2, 6)
        assert 2 <= r["total"] <= 12
        assert len(r["rolls"]) == 2
        assert all(1 <= x <= 6 for x in r["rolls"])


def test_roll_single():
    for _ in range(100):
        r = roll(1, 20)
        assert 1 <= r["total"] <= 20


# ── roll_attack ───────────────────────────────────────────────────────────────

def test_nat20_always_hits_and_crits():
    import random
    random.seed(None)
    # Force nat 20 by monkeypatching
    import engine.dice as dice_mod
    orig = random.randint
    random.randint = lambda a, b: 20
    try:
        r = dice_mod.roll_attack(to_hit=0, target_ac=30)
        assert r["hit"] is True
        assert r["crit"] is True
        assert r["d20"] == 20
    finally:
        random.randint = orig


def test_nat1_always_misses():
    import random
    import engine.dice as dice_mod
    orig = random.randint
    random.randint = lambda a, b: 1
    try:
        r = dice_mod.roll_attack(to_hit=100, target_ac=1)
        assert r["hit"] is False
        assert r["crit"] is False
    finally:
        random.randint = orig


def test_normal_hit():
    import random
    import engine.dice as dice_mod
    orig = random.randint
    random.randint = lambda a, b: 15
    try:
        r = dice_mod.roll_attack(to_hit=5, target_ac=18)
        # 15 + 5 = 20 >= 18 → hit
        assert r["hit"] is True
        assert r["crit"] is False
        assert r["total"] == 20
    finally:
        random.randint = orig


def test_normal_miss():
    import random
    import engine.dice as dice_mod
    orig = random.randint
    random.randint = lambda a, b: 5
    try:
        r = dice_mod.roll_attack(to_hit=3, target_ac=20)
        # 5 + 3 = 8 < 20 → miss
        assert r["hit"] is False
    finally:
        random.randint = orig


def test_advantage_takes_higher():
    import random
    import engine.dice as dice_mod
    seq = iter([8, 15])
    orig = random.randint
    random.randint = lambda a, b: next(seq)
    try:
        r = dice_mod.roll_attack(to_hit=0, target_ac=1, adv=True)
        assert r["d20"] == 15
    finally:
        random.randint = orig


def test_disadvantage_takes_lower():
    import random
    import engine.dice as dice_mod
    seq = iter([15, 8])
    orig = random.randint
    random.randint = lambda a, b: next(seq)
    try:
        r = dice_mod.roll_attack(to_hit=0, target_ac=1, dis=True)
        assert r["d20"] == 8
    finally:
        random.randint = orig


# ── roll_damage ───────────────────────────────────────────────────────────────

def test_crit_doubles_dice_not_bonus():
    import random
    import engine.dice as dice_mod
    # All dice roll 4; so 2d6+3 normal = 8+3=11, crit 4d6+3 = 16+3=19
    orig = random.randint
    random.randint = lambda a, b: 4
    try:
        dd = [DamageDie(n=2, die=6, bonus=3, damage_type="slashing")]
        normal = dice_mod.roll_damage(dd, crit=False)
        crit_r = dice_mod.roll_damage(dd, crit=True)
        assert normal["grand_total"] == 11   # 2*4 + 3
        assert crit_r["grand_total"] == 19   # 4*4 + 3
        assert len(crit_r["clauses"][0]["rolls"]) == 4
    finally:
        random.randint = orig


def test_multi_clause_damage():
    import random
    import engine.dice as dice_mod
    orig = random.randint
    random.randint = lambda a, b: 3
    try:
        dds = [
            DamageDie(n=1, die=8, bonus=3, damage_type="slashing"),
            DamageDie(n=2, die=6, bonus=0, damage_type="fire"),
        ]
        r = dice_mod.roll_damage(dds)
        # clause 1: 3+3=6, clause 2: 3+3=6 → 12
        assert r["grand_total"] == 12
        assert len(r["clauses"]) == 2
    finally:
        random.randint = orig


def test_damage_floor_never_below_bonus():
    dd = [DamageDie(n=1, die=6, bonus=10, damage_type="")]
    r = roll_damage(dd)
    # Even minimum roll (1) + 10 bonus = 11
    assert r["grand_total"] >= 11


# ── roll_batch ────────────────────────────────────────────────────────────────

def test_batch_count():
    r = roll_batch(count=5, to_hit=5, target_ac=12)
    assert len(r["attacks"]) == 5
    assert r["count"] == 5
    assert 0 <= r["hits"] <= 5


# ── mob_hits ──────────────────────────────────────────────────────────────────

def test_mob_required_roll_clamp():
    # AC 10, to_hit +15 → required = max(1, 10-15) = 1
    r = mob_hits(attackers=20, to_hit=15, target_ac=10)
    assert r["required_roll"] == 1
    assert r["per_hit"] == 1
    assert r["hits"] == 20

    # AC 30, to_hit 0 → required = min(20, 30) = 20
    r2 = mob_hits(attackers=20, to_hit=0, target_ac=30)
    assert r2["required_roll"] == 20
    assert r2["per_hit"] == 20
    assert r2["hits"] == 1


def test_mob_override():
    r = mob_hits(attackers=10, to_hit=5, target_ac=15, override=7)
    assert r["hits"] == 7
    assert r["override_used"] is True


def test_mob_no_override():
    r = mob_hits(attackers=10, to_hit=5, target_ac=15)
    assert r["override_used"] is False
    assert r["hits"] == r["computed_hits"]


# ── avg_damage ────────────────────────────────────────────────────────────────

def test_avg_damage():
    dd = [DamageDie(n=2, die=8, bonus=4, damage_type="fire")]
    # avg of 2d8 = 2 * 4.5 = 9; +4 = 13
    assert avg_damage(dd) == pytest.approx(13.0)
