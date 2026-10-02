"""Compute every table of the v42 workbook as plain Python data (no openpyxl here)."""
import re
from . import jalali as J, engine as E, snapshots as S, load_v41 as L, curated as C
from .names import norm, pid

BASE, END = S.BASE, S.END
EXPLICIT_FY1403_END = {'شتران': 'گزارش SRC_SHTR_1403_001 دوره منتهی به 1403/12/30',
                       'فخوز': 'گزارش SRC_FKHZ_1404_002 منتهی به 1403/12/30',
                       'وپاسار': 'گزارش فعالیت 1396276 منتهی به 1403/12/30'}
TIER_LABEL = {
    'A_FULL': 'دو فهرست اسمی کامل',
    'A_SENS': 'دو فهرست اسمی کامل با حساسیت ثبت شده',
    'A_COND': 'فهرست پایان مشروط',
    'H_HOLDER': 'فقط سطح شخص حقوقی؛ نماینده ها نامشخص',
    'B_AGG': 'فقط اعداد تجمیعی ارثی؛ فهرست اسمی در دسترس نیست',
    'B_CONTRA': 'عدد تجمیعی ارثی متعارض با شواهد v41',
    'C_NONE': 'بدون داده کافی',
}
TIER_A = {'تاپیکو': 'A_FULL', 'کچاد': 'A_FULL', 'خساپا': 'A_FULL', 'خودرو': 'A_FULL', 'شپدیس': 'A_FULL',
          'جم': 'A_SENS', 'پارسان': 'A_SENS', 'وتجارت': 'A_COND'}
DIRECT_STRUCT = {'وتجارت', 'وبملت', 'وبصادر'}      # individual directors, no legal seats
STRUCT_UNDETERMINED = {'وپاسار'}


def d10(s):
    return (s or '')[:10]


def fy_info(sc):
    start = sc['fy_start']
    fye = sc['fye']
    flag = ''
    if start == '1403/12/30':
        fy3 = '1403/12/30'
        basis = 'پایان کامل سال تقویمی 1403 (کبیسه)'
        flag = 'ستون شروع سال مالی در v41 برابر 1403/12/30 است ولی شروع سال 1404 روز 1404/01/01 است؛ مقدار v41 دست نخورد'
    else:
        fy3 = J.add_days(start, -1)
        basis = 'یک روز پیش از شروع FY1404 در v41'
    src = EXPLICIT_FY1403_END.get(sc['sym'], 'استنباط تقویمی؛ سند صریح در این فایل ثبت نشده')
    return fy3, basis, src, flag


def registry_by_company(reg):
    out = {}
    for r in reg:
        if r['type'].startswith('اطلاعیه'):
            out.setdefault(r['sym'], []).append(r)
    for k in out:
        out[k].sort(key=lambda r: (d10(r['pub']), r['tracing']))
    return out


def build_model(wb, wbv):
    m = {}
    scope = L.read_scope(wb); master = L.read_master(wb); seat = L.read_seat_recon(wb)
    audit = L.read_audit(wb); ten = L.read_tenure(wb); roster = L.read_roster_v41(wb)
    reg = L.read_registry(wb); cr1a = L.read_cr1a(wb)
    m.update(scope=scope, master=master, seat=seat, audit=audit, tenure=ten, roster=roster,
             registry=reg, cr1a=cr1a)
    syms = [s['sym'] for s in scope]
    scd = {s['sym']: s for s in scope}
    m['syms'] = syms; m['scd'] = scd
    notices = registry_by_company(reg)
    m['notices'] = notices

    # ---- tiers
    tier = {}
    for s in syms:
        if s in TIER_A:
            tier[s] = TIER_A[s]
        elif s == 'حکشتی':
            tier[s] = 'H_HOLDER'
        elif s == 'وپاسار':
            tier[s] = 'C_NONE'
        elif s == 'فملی':
            tier[s] = 'B_CONTRA'
        else:
            tier[s] = 'B_AGG'
    m['tier'] = tier

    # ---- time basis
    tb = {}
    for s in scope:
        sym = s['sym']
        fy3, fy3basis, fy3src, flag = fy_info(s)
        d1 = J.days_between(BASE, fy3)
        d2 = J.days_between(END, s['fye'])
        if s['fye'] == END and fy3 == '1403/12/30':
            rel = 'پایان یکسان؛ مبنا با یک روز اختلاف (1403/12/29 در برابر 1403/12/30)'
        elif s['fye'] == END:
            rel = 'پایان یکسان؛ مبنا متفاوت'
        else:
            rel = 'پایان و مبنا متفاوت'
        ns = notices.get(sym, [])
        between = [n for n in ns if BASE < d10(n['pub']) <= END]
        after = [n for n in ns if d10(n['pub']) > END]
        gap = [n for n in ns if d10(n['pub']) in ('1403/12/30',)]
        before = [n for n in ns if d10(n['pub']) <= BASE]
        upto_end = [n for n in ns if d10(n['pub']) <= END]
        tb[sym] = dict(fy3=fy3, fy3basis=fy3basis, fy3src=fy3src, flag=flag, d1=d1, d2=d2, rel=rel,
                       n_between=len(between), n_after=len(after), n_gap=len(gap), n_all=len(ns),
                       base_notice=(before[-1] if before else None), end_notice=(upto_end[-1] if upto_end else None),
                       valid=all(J.is_valid(x) for x in (BASE, END, s['fye'], fy3, s['fy_start'] if s['fy_start'] != '1403/12/30' else '1404/01/01')))
    m['tb'] = tb

    # ---- Tier A snapshots
    snaps = S.build(seat)
    m['snaps'] = snaps
    flows, seatrows, legacy, structure, chairs = {}, {}, {}, {}, {}
    for sym, (b, e, meta) in snaps.items():
        flows[sym] = E.person_flows(b, e)
        seatrows[sym] = E.seat_changes(b, e)
        legacy[sym] = E.legacy_reproduce(b, e)
        structure[sym] = E.seat_structure(b, e)
        chairs[sym] = E.chair_change(b, e)
    m.update(flows=flows, seatrows=seatrows, legacy=legacy, structure=structure, chairs=chairs)

    # ---- Hokeshti holder level (from v38 audit row text in the workbook)
    hk = [('مادر تخصصی بازرگانی دولتی ایران', 'تأمین اندیش رفاه پردیس'),
          ('صندوق بازنشستگی کشوری', 'آتیه صبا'),
          ('سرمایه گذاری استان گیلان', 'سیستان و بلوچستان'),
          ('سرمایه گذاری تأمین اجتماعی', 'سرمایه گذاری تأمین اجتماعی'),
          ('سرمایه گذاری صبا تأمین', 'سرمایه گذاری صبا تأمین')]
    hb, he = [], []
    for a, b2 in hk:
        key = norm(a if a == b2 else f'{a} → {b2}')
        hb.append(E.Seat(norm(a), None, E.UNNAMED, 'نامشخص', key, 'فاقد نماینده', a, '1310210'))
        he.append(E.Seat(norm(b2), None, E.UNNAMED, 'نامشخص', key, 'فاقد نماینده', b2, '1433770'))
    m['hk_base'] = E.Snapshot('حکشتی', BASE, hb, False)
    m['hk_end'] = E.Snapshot('حکشتی', END, he, False)
    m['hk_rows'] = E.seat_changes(m['hk_base'], m['hk_end'])
    return m


def tenure_relation(t):
    """Relation of a FYE-roster member to the common base cutoff, from v41 tenure dates.
    Returns (code, text).  Only dates already valid in v41; nothing invented."""
    cls = t['cls']; st = t['start']; lo = t['lower']; up = t['upper']
    if cls == 'EXACT':
        if st > BASE:
            return 'ENTRY_AFTER_BASE', 'ورود پس از مبنا (شروع دقیق پس از 1403/12/29)'
        return 'PRESENT_AT_BASE_BY_CONTINUITY', 'حاضر در مبنا بر پایه پیوستگی (شروع دقیق تا 1403/12/29)'
    if cls == 'INTERVAL':
        if lo and lo > BASE:
            return 'ENTRY_AFTER_BASE', 'ورود پس از مبنا (کران پایین پس از 1403/12/29)'
        if up and up <= BASE:
            return 'PRESENT_AT_BASE_BY_CONTINUITY', 'حاضر در مبنا بر پایه پیوستگی (کران بالا تا 1403/12/29)'
        return 'BASE_STATUS_UNDETERMINED', 'نامشخص: بازه شروع مبنا را دربر می گیرد'
    if cls in ('MIN_CONTINUOUS', 'CONFLICT_MIN'):
        a = up or lo or st
        if a and a <= BASE:
            return 'PRESENT_AT_BASE_BY_CONTINUITY', 'حاضر در مبنا بر پایه حداقل پیوستگی (لنگر تا 1403/12/29)'
        return 'BASE_STATUS_UNDETERMINED', 'نامشخص: لنگر پس از مبنا و شروع واقعی معلوم نیست'
    return 'BASE_STATUS_UNDETERMINED', 'نامشخص: تصدی ناشناخته'


def roster_rows(m):
    """FYE1404 roster (v41) joined with tenure evidence."""
    ten_by = {}
    for t in m['tenure']:
        if t['kind'] == 'MEMBER':
            ten_by[t['member_id']] = t
    rows = []
    for r in m['roster']:
        t = ten_by.get(r['id'])
        rel = tenure_relation(t) if t else ('NO_TENURE_ROW', 'ردیف تصدی ندارد')
        rows.append(dict(r=r, t=t, rel=rel))
    return rows


def entrants_check(m, rrows):
    """Compare master P with tenure-based entrants.  Returns per-company dict."""
    out = {}
    for sym in m['syms']:
        rs = [x for x in rrows if x['r']['sym'] == sym and x['t'] is not None]
        ent = sum(1 for x in rs if x['rel'][0] == 'ENTRY_AFTER_BASE')
        und = sum(1 for x in rs if x['rel'][0] == 'BASE_STATUS_UNDETERMINED')
        pres = sum(1 for x in rs if x['rel'][0] == 'PRESENT_AT_BASE_BY_CONTINUITY')
        P = m['master'][sym]['P']
        same_end = m['scd'][sym]['fye'] == END
        if not rs:
            flag = 'فهرست v41 ندارد'
        elif P is None:
            flag = 'مقدار مستر ندارد'
        elif same_end:
            flag = 'هم راستا' if ent <= P <= ent + und else 'متعارض'
        else:
            flag = 'هم راستا (کف تا افق FYE)' if ent <= P else 'متعارض'
        out[sym] = dict(entrants=ent, undetermined=und, present=pres, n=len(rs), P=P, same_end=same_end, flag=flag)
    return out
