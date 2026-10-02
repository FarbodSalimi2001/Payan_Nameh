"""Person-flow / seat-change / measure engine for the two-snapshot comparison.

Design rules (from the stage-3 brief):
  * a person is identified by normalised name inside one company (no fuzzy merge);
  * absence from a list is an exit only if BOTH lists are complete;
  * 'no representative named' is not an exit, not a vacancy and not zero change;
  * seat pairing is done only through an explicit seat_key (legal holder or a
    documented holder replacement); direct individual seats are never paired;
  * unique-person measures are separate from the legacy seat-based REF05 family.
"""
from dataclasses import dataclass, field
from typing import Optional
from .names import norm

NAMED = 'NAMED'
VACANT_DOC = 'VACANT_DOCUMENTED'      # vacancy documented (representative removed, no successor)
UNNAMED = 'UNNAMED_UNVERIFIED'        # 'no representative' text, vacancy not established
UNK = 'UNKNOWN'


@dataclass
class Seat:
    holder: Optional[str]              # normalised legal holder; None for a direct individual member
    rep: Optional[str]                 # normalised person name or None
    rep_state: str = NAMED
    role: str = ''
    seat_key: Optional[str] = None     # None -> cannot be paired across snapshots
    raw_rep: str = ''
    raw_holder: str = ''
    src: str = ''


@dataclass
class Snapshot:
    company: str
    cutoff: str
    seats: list = field(default_factory=list)
    complete: bool = False
    caveat: str = ''

    def persons(self):
        seen, out = set(), []
        for s in self.seats:
            if s.rep and s.rep_state == NAMED and s.rep not in seen:
                seen.add(s.rep)
                out.append(s.rep)
        return out

    def duplicate_persons(self):
        names = [s.rep for s in self.seats if s.rep and s.rep_state == NAMED]
        return sorted({n for n in names if names.count(n) > 1})

    def chair(self):
        c = [s.rep for s in self.seats if s.rep and s.rep_state == NAMED and is_chair(s.role)]
        return c[0] if len(c) == 1 else None

    def has_unnamed(self):
        return any(s.rep_state in (UNNAMED, UNK) for s in self.seats)


def role_title(role):
    t = norm(role).split('(')[0].strip()
    return t


def is_chair(role):
    t = role_title(role)
    return t.startswith('رئیس') and 'نایب' not in t


def is_ceo(role):
    return 'مدیرعامل' in role_title(role).replace(' ', '') or 'مدیر عامل' in role_title(role)


def is_executive(role):
    r = norm(role)
    return ('موظف' in r) and ('غیر موظف' not in r) and ('غیرموظف' not in r)


# ---------------------------------------------------------------- person flows
def person_flows(base: Snapshot, end: Snapshot):
    """Unique-person comparison of two snapshots of the same company.

    Returns dict with per-person records and counts.  Rates are produced only
    when both snapshots are flagged complete and have no unnamed seat."""
    assert base.company == end.company, 'comparison must stay inside one company'
    bp, ep = base.persons(), end.persons()
    bset, eset = set(bp), set(ep)
    both_complete = base.complete and end.complete and not base.has_unnamed() and not end.has_unnamed()
    recs = []
    for n in bp:
        if n in eset:
            recs.append((n, 'PERSISTED'))
        else:
            recs.append((n, 'EXIT_BETWEEN_SNAPSHOTS' if both_complete else 'ABSENT_AT_END_UNDETERMINED'))
    for n in ep:
        if n not in bset:
            recs.append((n, 'ENTRY_BETWEEN_SNAPSHOTS' if both_complete else 'PRESENT_AT_END_NOT_IN_BASE_UNDETERMINED'))
    res = dict(records=recs, n_base=len(bp), n_end=len(ep), complete=both_complete,
               persisted=sum(1 for _, s in recs if s == 'PERSISTED'),
               exits=sum(1 for _, s in recs if s == 'EXIT_BETWEEN_SNAPSHOTS'),
               entries=sum(1 for _, s in recs if s == 'ENTRY_BETWEEN_SNAPSHOTS'))
    if both_complete and res['n_base'] > 0 and res['n_end'] > 0:
        res['exit_rate'] = (res['exits'], res['n_base'])
        res['entry_share'] = (res['entries'], res['n_end'])
        res['persistence_rate'] = (res['persisted'], res['n_base'])
        res['balance_ok'] = (res['n_base'] - res['exits'] + res['entries'] == res['n_end'])
    else:
        res['exit_rate'] = res['entry_share'] = res['persistence_rate'] = None
        res['balance_ok'] = None
    return res


# ----------------------------------------------------------------- seat changes
SEAT_TYPES = {
    'STABLE': 'ثبات شخص حقوقی و نماینده',
    'REP_CHANGE': 'تغییر نماینده با ثبات شخص حقوقی',
    'HOLDER_CHANGE_SAME_PERSON': 'تغییر شخص حقوقی با ثبات فرد',
    'BOTH_CHANGE': 'تغییر همزمان شخص حقوقی و فرد',
    'PERSON_SWAP': 'جابجایی همان افراد بین کرسی ها',
    'ROLE_ONLY': 'تغییر سمت بدون تغییر عضویت',
    'FILLED_VACANT': 'پر شدن کرسی خالی مستند',
    'VACATED': 'خالی شدن مستند کرسی',
    'SEAT_ADDED_REMOVED': 'ایجاد یا حذف کرسی',
    'UNMATCHED': 'تطبیق نامشخص',
}


def _seat_map(snap):
    m = {}
    for s in snap.seats:
        if s.seat_key is not None:
            m.setdefault(s.seat_key, []).append(s)
    return m


def seat_changes(base: Snapshot, end: Snapshot):
    """Classify every seat that can be paired by seat_key.  Direct individual
    seats (seat_key None) are returned separately and are never paired."""
    bm, em = _seat_map(base), _seat_map(end)
    bset, eset = set(base.persons()), set(end.persons())
    rows = []
    for key in list(bm) + [k for k in em if k not in bm]:
        b = bm.get(key, [None])[0]
        e = em.get(key, [None])[0]
        if b is None or e is None:
            rows.append(dict(key=key, base=b, end=e, type='SEAT_ADDED_REMOVED',
                             moved_in=False, moved_out=False))
            continue
        if b.rep_state not in (NAMED, VACANT_DOC) or e.rep_state not in (NAMED, VACANT_DOC):
            rows.append(dict(key=key, base=b, end=e, type='UNMATCHED', moved_in=False, moved_out=False))
            continue
        holder_same = (b.holder == e.holder)
        if b.rep_state == VACANT_DOC and e.rep_state == NAMED:
            t = 'FILLED_VACANT'
        elif b.rep_state == NAMED and e.rep_state == VACANT_DOC:
            t = 'VACATED'
        elif b.rep_state == VACANT_DOC and e.rep_state == VACANT_DOC:
            t = 'STABLE' if holder_same else 'HOLDER_CHANGE_SAME_PERSON'
        else:
            rep_same = (b.rep == e.rep)
            if holder_same and rep_same:
                t = 'STABLE' if role_title(b.role) == role_title(e.role) else 'ROLE_ONLY'
            elif holder_same and not rep_same:
                t = 'REP_CHANGE'
            elif (not holder_same) and rep_same:
                t = 'HOLDER_CHANGE_SAME_PERSON'
            else:
                t = 'BOTH_CHANGE'
        moved_in = bool(e.rep and e.rep in bset and (not b.rep or e.rep != b.rep))
        moved_out = bool(b.rep and b.rep in eset and (not e.rep or e.rep != b.rep))
        rows.append(dict(key=key, base=b, end=e, type=t, moved_in=moved_in, moved_out=moved_out))
    # pure swap: both sides of a changed seat are persons present in both snapshots
    for r in rows:
        if r['type'] == 'REP_CHANGE' and r['moved_in'] and r['moved_out']:
            r['type'] = 'PERSON_SWAP'
    return rows


def direct_seat_count(snap):
    return sum(1 for s in snap.seats if s.seat_key is None)


# --------------------------------------------------------------- legacy family
def legacy_reproduce(base: Snapshot, end: Snapshot):
    """Re-derive the legacy REF05/REF08/ROB_* numerators and denominators with
    the definitions documented in 'Turnover نهایی'!Y (seat based):
      G/H  REF05: comparable seats whose named person differs / comparable seats
      J/K  REF08: comparable seats with the SAME legal holder and a different
                  person / comparable seats with the same legal holder
      M/N  ROB_LEGAL: seats whose legal holder differs / seats present in both
      P/Q  ROB_NEW: end persons not in base / end persons
      S/T  ROB_STABILITY: comparable seats with the same person / comparable seats
    Direct individual seats (no seat_key) are compared at person level:
      G = base persons absent at end, H = base persons, S = persons present in both."""
    out = dict(G=None, H=None, J=None, K=None, M=None, N=None, P=None, Q=None, S=None, T=None)
    bset, eset = set(base.persons()), set(end.persons())
    out['P'] = len([n for n in end.persons() if n not in bset])
    out['Q'] = len(end.persons())
    rows = seat_changes(base, end)
    if direct_seat_count(base) or direct_seat_count(end):
        out['G'] = len([n for n in base.persons() if n not in eset])
        out['H'] = len(base.persons())
        out['S'] = out['H'] - out['G']
        out['T'] = out['H']
        return out
    comp = [r for r in rows if r['base'] and r['end'] and r['base'].rep_state == NAMED and r['end'].rep_state == NAMED]
    out['H'] = len(comp)
    out['G'] = sum(1 for r in comp if r['base'].rep != r['end'].rep)
    stable_holder = [r for r in comp if r['base'].holder == r['end'].holder]
    out['K'] = len(stable_holder)
    out['J'] = sum(1 for r in stable_holder if r['base'].rep != r['end'].rep)
    both = [r for r in rows if r['base'] and r['end']]
    out['N'] = len(both)
    out['M'] = sum(1 for r in both if r['base'].holder != r['end'].holder)
    out['S'] = out['H'] - out['G']
    out['T'] = out['H']
    return out


def seat_structure(base: Snapshot, end: Snapshot):
    """Counts the brief asks to be kept separate for REF08-type measures."""
    rows = seat_changes(base, end)
    both = [r for r in rows if r['base'] and r['end']]
    stable = [r for r in both if r['base'].holder == r['end'].holder]
    known = [r for r in stable if r['base'].rep_state == NAMED and r['end'].rep_state == NAMED]
    unknown = [r for r in stable if r['base'].rep_state in (UNNAMED, UNK) or r['end'].rep_state in (UNNAMED, UNK)]
    vacant = [r for r in stable if r not in known and r not in unknown]
    return dict(seats_both=len(both), stable_holder_seats=len(stable), stable_known_both=len(known),
                stable_unknown_rep=len(unknown), stable_vacant_documented=len(vacant),
                ref08_changed=sum(1 for r in known if r['base'].rep != r['end'].rep))


# ------------------------------------------------------------------- chair / events
def chair_change(base: Snapshot, end: Snapshot):
    b, e = base.chair(), end.chair()
    if b is None or e is None:
        return None, b, e
    return (1 if b != e else 0), b, e


def person_events(events):
    """events: iterable of (person, 'ENTRY'|'EXIT', date).  Counts documented events
    inside a window; a person who leaves and returns contributes two events even
    though two-endpoint comparison shows him persisted."""
    return {'entries': sum(1 for _, k, _d in events if k == 'ENTRY'),
            'exits': sum(1 for _, k, _d in events if k == 'EXIT')}
