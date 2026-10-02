"""Build comparable two-cutoff snapshots from the v38 seat reconstruction rows."""
import re
from .names import norm
from . import engine as E

BASE = '1403/12/29'
END = '1404/12/29'

# Completeness judgement per company (curated from the v38 audit text that sits in
# the v41 workbook; the underlying notices could not be re-opened in this stage).
# status: FULL | FULL_SENSITIVITY | CONDITIONAL
COMPLETENESS = {
    'تاپیکو': ('FULL', 'FULL', ''),
    'کچاد': ('FULL', 'FULL', 'کرسی خوزستان در مبنا خالی مستند: نماینده قبلی ۱۴۰۳/۰۹/۰۴ حذف و جانشین تا ۱۴۰۴/۰۵/۲۰ نیامد (متن v38)'),
    'پارسان': ('FULL', 'FULL_SENSITIVITY', 'مجمع ۱۴۰۴/۱۰/۲۳ دو شخص حقوقی را حذف و دو شخص حقوقی تازه وارد کرد؛ v38 آن را تا Cutoff عملیاتی ندانست و نمایندگان تازه فقط در جلسه ۱۴۰۵/۰۲/۰۸ ثبت شدند؛ فهرست پایان براساس اطلاعیه پیش از مجمع'),
    'جم': ('FULL', 'FULL_SENSITIVITY', 'فهرست پایان از اطلاعیه ۱۵۲۷۳۸۹ با جلسه موثر ۱۴۰۴/۱۲/۲۵ و انتشار ۱۴۰۵/۰۲/۲۹ (پس از Cutoff)؛ تاریخ موثر صریح است اما با مبنای انتشار معینی و تغییر رئیس حذف می شوند'),
    'خساپا': ('FULL', 'FULL', ''),
    'خودرو': ('FULL', 'FULL', ''),
    'شپدیس': ('FULL', 'FULL', 'کرسی هامون سپاهان در مبنا خالی مستند: نماینده قبلی ۱۴۰۳/۱۱/۲۹ حذف و تا ۱۴۰۴/۰۸/۱۷ جانشین نیامد (متن v38)'),
    'وتجارت': ('FULL', 'CONDITIONAL', 'فهرست پایان چهار نفره از اصلاحیه ۱۴۴۴۱۵۴ و جدول FY1404 است؛ اطلاعیه اولیه ۱۴۴۳۰۷۸ با تاریخ جلسه دیرتر هنوز پنج نفر دارد؛ خالی بودن کرسی پنجم مستند نیست'),
}

VACANT_NOTE = {
    ('کچاد', 'سرمایه گذاری استان خوزستان'): 'حذف نماینده قبلی ۱۴۰۳/۰۹/۰۴ و نبود جانشین تا ۱۴۰۴/۰۵/۲۰',
    ('شپدیس', 'سرمایه گذاری هامون سپاهان'): 'حذف نماینده قبلی ۱۴۰۳/۱۱/۲۹ و نبود جانشین تا ۱۴۰۴/۰۸/۱۷',
}


def _cutoff(label):
    m = re.search(r'(\d{4}/\d{2}/\d{2})', label or '')
    return m.group(1) if m else None


def build(seat_rows):
    """Returns {company: (base_snapshot, end_snapshot, meta)} for companies that have
    a named two-cutoff reconstruction."""
    by = {}
    for r in seat_rows:
        c = _cutoff(r['snapshot'])
        if c not in (BASE, END):
            continue
        by.setdefault(r['sym'], {}).setdefault(c, []).append(r)
    out = {}
    for sym, d in by.items():
        if BASE not in d or END not in d:
            continue
        comp_b, comp_e, caveat = COMPLETENESS[sym][0], COMPLETENESS[sym][1], COMPLETENESS[sym][2]
        snaps = {}
        for cut in (BASE, END):
            seats = []
            for r in d[cut]:
                holder_raw = r['holder'] or ''
                rep_raw = r['rep'] or ''
                direct = 'عضو حقیقی مستقیم' in holder_raw
                src = f"{r['tracing']}"
                if direct:
                    if 'خارج شده' in rep_raw:
                        continue                      # exit marker, not a member
                    seats.append(E.Seat(None, norm(rep_raw), E.NAMED, r['role'] or '', None, rep_raw, holder_raw, src))
                    continue
                if '→' in holder_raw:
                    a, b = [x.strip() for x in holder_raw.split('→')]
                    key = norm(holder_raw)
                    hold = norm(a if cut == BASE else b)
                else:
                    key = norm(holder_raw)
                    hold = key
                if 'کرسی بدون نماینده' in rep_raw:
                    st = E.VACANT_DOC if (sym, key) in VACANT_NOTE else E.UNNAMED
                    seats.append(E.Seat(hold, None, st, r['role'] or '', key, rep_raw, holder_raw, src))
                else:
                    seats.append(E.Seat(hold, norm(rep_raw), E.NAMED, r['role'] or '', key, rep_raw, holder_raw, src))
            snaps[cut] = E.Snapshot(sym, cut, seats)
        snaps[BASE].complete = comp_b.startswith('FULL')
        snaps[END].complete = comp_e.startswith('FULL')     # CONDITIONAL -> not complete
        snaps[END].caveat = caveat
        out[sym] = (snaps[BASE], snaps[END], dict(base=comp_b, end=comp_e, caveat=caveat))
    return out
