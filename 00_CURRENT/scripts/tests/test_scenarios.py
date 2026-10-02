"""Stage-3 acceptance scenarios (synthetic, no workbook needed).  Run: python -m pytest tests"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from bsi_v42 import engine as E
from bsi_v42.names import norm


def snap(cut, seats, complete=True):
    return E.Snapshot('X', cut, seats, complete)


def S(key, rep, role='عضو (غیرموظف)', holder=None, state=E.NAMED):
    return E.Seat(holder or key, norm(rep) if rep else None, state, role, key)


def test_same_people_swap_between_two_stable_legal_seats():
    b = snap('1403/12/29', [S('A', 'علی'), S('B', 'رضا'), S('C', 'مینا')])
    e = snap('1404/12/29', [S('A', 'رضا'), S('B', 'علی'), S('C', 'مینا')])
    pf = E.person_flows(b, e)
    assert (pf['exits'], pf['entries'], pf['persisted']) == (0, 0, 3)
    rows = {r['key']: r for r in E.seat_changes(b, e)}
    assert rows['A']['type'] == 'PERSON_SWAP' and rows['B']['type'] == 'PERSON_SWAP'
    assert E.legacy_reproduce(b, e)['J'] == 2      # rep-change positive while person flow is zero


def test_new_person_replaces_rep_of_same_legal_holder():
    b = snap('b', [S('A', 'علی'), S('B', 'رضا')]); e = snap('e', [S('A', 'سارا'), S('B', 'رضا')])
    pf = E.person_flows(b, e)
    assert (pf['exits'], pf['entries']) == (1, 1)
    r = {x['key']: x for x in E.seat_changes(b, e)}
    assert r['A']['type'] == 'REP_CHANGE' and r['A']['base'].holder == r['A']['end'].holder


def test_holder_changes_person_stays():
    b = snap('b', [E.Seat('H1', norm('علی'), E.NAMED, 'عضو', 'SEAT1')])
    e = snap('e', [E.Seat('H2', norm('علی'), E.NAMED, 'عضو', 'SEAT1')])
    pf = E.person_flows(b, e)
    assert (pf['exits'], pf['entries'], pf['persisted']) == (0, 0, 1)
    assert E.seat_changes(b, e)[0]['type'] == 'HOLDER_CHANGE_SAME_PERSON'
    assert E.legacy_reproduce(b, e)['M'] == 1


def test_exit_and_return_inside_window():
    b = snap('b', [S('A', 'علی')]); e = snap('e', [S('A', 'علی')])
    assert E.person_flows(b, e)['persisted'] == 1
    ev = [('علی', 'EXIT', '1404/03/01'), ('علی', 'ENTRY', '1404/09/01')]
    assert E.person_events(ev) == {'entries': 1, 'exits': 1}   # endpoint view: 0/0, event view: 1/1


def test_role_only_and_reelection_do_not_create_flow():
    b = snap('b', [S('A', 'علی', 'رئیس هیئت مدیره (غیرموظف)')])
    e = snap('e', [S('A', 'علی', 'عضو (غیرموظف)')])
    pf = E.person_flows(b, e)
    assert pf['exits'] == 0 and pf['entries'] == 0
    assert E.seat_changes(b, e)[0]['type'] == 'ROLE_ONLY'
    assert E.chair_change(b, e)[0] is None          # end chair unknown -> no inference


def test_unknown_representative_is_not_zero_change_nor_exit():
    b = snap('b', [S('A', 'علی'), S('B', 'رضا')])
    e = snap('e', [S('A', 'علی'), E.Seat('B', None, E.UNNAMED, 'عضو', 'B')])
    pf = E.person_flows(b, e)
    assert pf['complete'] is False and pf['exit_rate'] is None
    assert [s for _, s in pf['records'] if s.startswith('ABSENT')] == ['ABSENT_AT_END_UNDETERMINED']
    row = {r['key']: r for r in E.seat_changes(b, e)}['B']
    assert row['type'] == 'UNMATCHED'
    ss = E.seat_structure(b, e)
    assert ss['stable_unknown_rep'] == 1 and ss['stable_known_both'] == 1


def test_complete_lists_identities():
    b = snap('b', [S('A', 'علی'), S('B', 'رضا'), S('C', 'مینا'), S('D', 'کامران')])
    e = snap('e', [S('A', 'علی'), S('B', 'سارا'), S('C', 'مینا'), S('D', 'نگار')])
    pf = E.person_flows(b, e)
    assert pf['balance_ok'] is True
    ex, nb = pf['exit_rate']; pe, _ = pf['persistence_rate']
    assert ex + pe == nb                              # complementary, same denominator


def test_names_not_merged_by_similarity():
    assert norm('محمد رضا کریمی') != norm('محمدرضا کریمی')   # no fuzzy merge
    assert norm('مرتضی ترک‌تبریزی') == norm('مرتضی ترک تبریزی')  # only ZWNJ/space unification
    assert norm('علي كريمي') == norm('علی کریمی')


def test_direct_seats_never_paired():
    b = snap('b', [E.Seat(None, norm('علی'), E.NAMED, 'عضو', None), E.Seat(None, norm('رضا'), E.NAMED, 'عضو', None)])
    e = snap('e', [E.Seat(None, norm('علی'), E.NAMED, 'عضو', None)], complete=False)
    assert E.seat_changes(b, e) == []
    pf = E.person_flows(b, e)
    assert pf['exits'] == 0 and pf['exit_rate'] is None   # incomplete end list -> not an exit


def test_new_persons_may_be_fewer_than_seat_changes():
    # general rule 'new >= seat changes' is false: swaps change seats without new people
    b = snap('b', [S('A', 'علی'), S('B', 'رضا')]); e = snap('e', [S('A', 'رضا'), S('B', 'علی')])
    assert E.legacy_reproduce(b, e)['G'] == 2 and E.legacy_reproduce(b, e)['P'] == 0
