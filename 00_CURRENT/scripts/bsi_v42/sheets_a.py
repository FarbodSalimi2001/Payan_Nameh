"""Scope, Snapshot Members, Person Flows, Flow Summary, Seat Changes."""
from . import engine as E, snapshots as S, prep as P, jalali as J
from .names import norm, pid
from .xl import Sheet, FX, FILL_WARN, FILL_BAD, FILL_OK

BASE, END = S.BASE, S.END
COMMON = 'پنجره مشترک (1403/12/29 تا 1404/12/29)'
FLOW_FA = {
    'PERSISTED': 'ماندگار: حاضر در هر دو مقطع',
    'EXIT_BETWEEN_SNAPSHOTS': 'خارج شده بین دو مقطع',
    'ENTRY_BETWEEN_SNAPSHOTS': 'واردشده بین دو مقطع',
    'ABSENT_AT_END_UNDETERMINED': 'غایب در پایان؛ خروج منوط به کامل بودن فهرست پایان (تعیین نشد)',
    'PRESENT_AT_END_NOT_IN_BASE_UNDETERMINED': 'حاضر در پایان و غایب در مبنا؛ ورود منوط به کامل بودن فهرست مبنا (تعیین نشد)',
    'FISCAL_ENTRY_AFTER_BASE': 'نمای مالی: ورود پس از مبنا طبق تصدی v41',
    'FISCAL_PRESENT_AT_BASE': 'نمای مالی: حاضر در مبنا طبق پیوستگی تصدی v41',
    'FISCAL_BASE_UNDETERMINED': 'نمای مالی: وضعیت در مبنا نامشخص',
    'FISCAL_NO_ROSTER': 'فهرست ندارد',
}
REL2FISCAL = {'ENTRY_AFTER_BASE': 'FISCAL_ENTRY_AFTER_BASE',
              'PRESENT_AT_BASE_BY_CONTINUITY': 'FISCAL_PRESENT_AT_BASE',
              'BASE_STATUS_UNDETERMINED': 'FISCAL_BASE_UNDETERMINED',
              'NO_TENURE_ROW': 'FISCAL_BASE_UNDETERMINED'}
COMP_FA = {'FULL': 'کامل', 'FULL_SENSITIVITY': 'کامل با حساسیت ثبت شده', 'CONDITIONAL': 'مشروط'}


def scope_sheet(wb, m):
    sh = Sheet(wb, 'Scope v42', 'دامنه v42: ۳۰ شرکت و دو مبنای زمانی',
               'نمای اصلی: پنجره مشترک 1403/12/29 تا 1404/12/29 برای مقایسه با اعداد قبلی. نمای مالی جدا: پایان دوره مالی 1403 و 1404 هر شرکت. دو نما در شمارش شرکت یکتا یا پوشش جمع نمی شوند. مقطع قراردادی 1403/12/29 آخرین روز واقعی سال 1403 نیست چون 1403 کبیسه است و روز 1403/12/30 وجود دارد؛ تاریخ قدیمی بی صدا تغییر نکرد.',
               ['شناسه', 'نماد', 'نام حقوقی', 'مقطع مبنای محاسبه قدیمی', 'مقطع پایان محاسبه قدیمی',
                'پایان دوره مالی 1403 (تقویم معتبر)', 'مبنای تعیین پایان FY1403', 'منبع پایان FY1403',
                'پایان دوره مالی 1404 (v41)', 'منبع و وضعیت تایید پایان FY1404',
                'تفاوت روز: پایان FY1403 منهای مبنای قدیمی', 'تفاوت روز: پایان FY1404 منهای پایان قدیمی',
                'رابطه دو مبنای زمانی', 'اعتبار تقویمی همه تاریخ ها', 'یادداشت تقویمی',
                'اطلاعیه ن۴۵ ثبت شده با انتشار بین دو مقطع', 'اطلاعیه ثبت شده با انتشار پس از پایان',
                'آخرین اطلاعیه ثبت شده تا مبنا (ردیابی و انتشار)', 'آخرین اطلاعیه ثبت شده تا پایان (ردیابی و انتشار)',
                'سطح داده در پنجره مشترک', 'نمای مالی: فهرست پایان FY1403', 'نمای مالی: فهرست پایان FY1404 (v41)',
                'نمای مالی قابل محاسبه', 'یادداشت'],
               widths={1: 7, 2: 10, 3: 30, 6: 14, 7: 24, 8: 28, 10: 30, 13: 30, 15: 40, 18: 22, 19: 22, 20: 26, 21: 26, 22: 26, 23: 14, 24: 40},
               raw_cols=(3,))
    for s in m['scope']:
        sym = s['sym']; tb = m['tb'][sym]
        bn, en = tb['base_notice'], tb['end_notice']
        nroster = sum(1 for x in m['roster'] if x['sym'] == sym and x['name'] and x['id'] != f"{s['id']}-NA")
        fye_status = f"{s['fye_status']}؛ {s['src_id']}"
        cal = tb['flag'] or ('تاریخ ها در تقویم معتبرند؛ مقاطع مشترک با پایان دوره مالی این شرکت یکی نیستند' if tb['rel'].startswith('پایان و مبنا') else 'تاریخ ها معتبرند؛ مبنای قراردادی یک روز پیش از پایان واقعی FY1403 است')
        lvl = P.TIER_LABEL[m['tier'][sym]]
        note = ''
        if sym == 'وپاسار':
            note = 'پایان FY1404 در v41 تایید نشده است؛ مقطع مالی و مشترک هر دو بدون فهرست'
        if sym in ('کچاد', 'کگل'):
            note = 'دوره 9 ماهه به دلیل تغییر سال مالی؛ پایان FY1403 پایان اسفند 1403 است'
        sh.add([s['id'], sym, s['legal'], BASE, END, tb['fy3'], tb['fy3basis'], tb['fy3src'], s['fye'], fye_status,
                tb['d1'], tb['d2'], tb['rel'], 'بله' if tb['valid'] else 'خیر', cal,
                tb['n_between'], tb['n_after'],
                f"{bn['tracing']} ({P.d10(bn['pub'])})" if bn else 'ندارد',
                f"{en['tracing']} ({P.d10(en['pub'])})" if en else 'ندارد',
                lvl, 'نامشخص: فهرست اسمی FY1403 در فایل نیست',
                f'{nroster} نام در v41؛ مقطع {s["fye"]}' if nroster else 'ندارد',
                'خیر', note], height=60)
    sh.autofilter()
    return sh


def _mem_status(seat, kind):
    if seat.rep_state == E.NAMED:
        return 'عضو اصلی فعال'
    if seat.rep_state == E.VACANT_DOC:
        return 'کرسی بدون نماینده؛ خالی بودن مستند (حذف نماینده قبلی و نبود جانشین)'
    return 'نماینده نامشخص؛ خالی بودن واقعی مستند نیست'


def members_sheet(wb, m):
    heads = ['شناسه ردیف', 'شناسه شرکت', 'نماد', 'مقطع', 'مبنای زمانی', 'نمای زمانی', 'شناسه پژوهشی شخص',
             'نام خام', 'نام یکسان سازی شده', 'مبنای هویت', 'وضعیت عضو', 'سمت', 'مدیرعامل', 'رئیس',
             'شخص حقوقی دارنده کرسی (خام)', 'شناسه شخص حقوقی', 'نماینده حقیقی', 'شناسه کرسی',
             'مبنای تطبیق کرسی', 'تعداد اعضای فعال ثبت شده در مستر', 'تعداد اشخاص شناسایی شده',
             'کامل یا ناقص', 'منبع اختصاصی', 'تاریخ منبع', 'نشانی منبع', 'اطمینان و تعارض', 'یادداشت']
    sh = Sheet(wb, 'Snapshot Members v42', 'فهرست اعضا و کرسی ها در هر مقطع',
               'نمای پنجره مشترک: هشت شرکت با فهرست اسمی دو مقطع (بازسازی سطح کرسی v38) و حکشتی در سطح شخص حقوقی. نمای مالی FYE1404: فهرست v41 با مقطع پایان دوره مالی هر شرکت؛ این نما با پنجره مشترک جمع نمی شود. هیچ مدیرعامل خارج از هیئت، مدعو، دبیر یا عضو کمیته به فهرست اعضا اضافه نشده است. شماره ردیف جدول شناسه کرسی نیست؛ شناسه کرسی فقط از نام دارنده حقوقی یا جایگزینی مستند ساخته شده است.',
               heads, widths={1: 9, 3: 9, 4: 12, 5: 18, 6: 18, 7: 12, 8: 20, 9: 20, 10: 24, 11: 28, 12: 22, 15: 28, 17: 20, 18: 26, 19: 26, 22: 16, 23: 16, 24: 12, 25: 20, 26: 28, 27: 40},
               freeze='D4', raw_cols=(8, 15, 17, 26, 27))
    n = 0
    sid = {s['sym']: s['id'] for s in m['scope']}
    master = m['master']
    # Tier A named snapshots
    for sym, (b, e, meta) in m['snaps'].items():
        seatrows = {(r['row']): r for r in m['seat']}
        for snap, cut, count, comp in ((b, BASE, master[sym]['H'], meta['base']), (e, END, master[sym]['Q'], meta['end'])):
            srcrows = [r for r in m['seat'] if r['sym'] == sym and S._cutoff(r['snapshot']) == cut]
            for r in srcrows:
                holder_raw = r['holder'] or ''
                rep_raw = r['rep'] or ''
                direct = 'عضو حقیقی مستقیم' in holder_raw
                marker = direct and 'خارج شده' in rep_raw
                n += 1
                nm = norm(rep_raw) if not marker and 'کرسی بدون نماینده' not in rep_raw else ''
                if marker:
                    st = 'نشانگر خروج؛ عضو فعال این مقطع نیست'
                elif 'کرسی بدون نماینده' in rep_raw:
                    st = 'کرسی بدون نماینده؛ خالی بودن مستند (حذف نماینده قبلی و نبود جانشین)'
                else:
                    st = 'عضو اصلی فعال'
                role = r['role'] or ''
                key = ('' if direct else norm(holder_raw))
                comp_fa = COMP_FA[comp]
                sh.add([f'SM{n:04d}', sid[sym], sym, cut, f'{cut} (مبنای قراردادی مشترک)', 'پنجره مشترک',
                        pid(sym, nm) if nm else '', rep_raw, nm, 'نام یکسان سازی شده در همان شرکت؛ کد ملی در این فایل نیست',
                        st, role, 'بله' if E.is_ceo(role) else 'خیر', 'بله' if (role and E.is_chair(role)) else 'خیر',
                        holder_raw, 'ندارد (فقط نام)' if not direct else 'ندارد (عضو حقیقی مستقیم)', '' if direct else rep_raw,
                        '' if direct else key,
                        'ممنوع: کرسی حقیقی مستقیم بدون شناسه پایدار' if direct else ('نام دارنده حقوقی؛ در صورت جایگزینی دارنده طبق اطلاعیه v38' if '→' in holder_raw else 'نام دارنده حقوقی'),
                        count, len(snap.persons()), comp_fa, str(r['tracing']), r['src_date'], r['url'], meta['caveat'] if cut == END or sym in ('کچاد', 'شپدیس') else '',
                        (r['note'] or '')[:300]], fill=(FILL_WARN if comp != 'FULL' else None), height=45)
    # Hokeshti holder-level
    for snap, cut, comp in ((m['hk_base'], BASE, 'مبنا'), (m['hk_end'], END, 'پایان')):
        for s in snap.seats:
            n += 1
            sh.add([f'SM{n:04d}', sid['حکشتی'], 'حکشتی', cut, f'{cut} (مبنای قراردادی مشترک)', 'پنجره مشترک', '', 'فاقد نماینده', '',
                    'نماینده حقیقی نامشخص', 'نماینده نامشخص؛ خالی بودن واقعی مستند نیست', '', '', '', s.raw_holder, 'ندارد (فقط نام)', '',
                    s.seat_key, 'جفت شدن دارنده ها طبق خوانش v38 از ستون قبلی و جدید اطلاعیه؛ بازبینی نشد',
                    '', 0, 'ناقص: نماینده ها نامشخص (دارندگان حقوقی کامل پنج کرسی)', s.src, '', '',
                    'سطح شخص حقوقی فقط', 'از اطلاعیه 1310210 و 1433770 طبق Turnover Resolution Audit'], fill=FILL_WARN, height=45)
    # FYE1404 roster (v41)
    rr = P.roster_rows(m)
    for x in rr:
        r = x['r']; t = x['t']
        sym = r['sym']
        n += 1
        if r['id'].endswith('-NA'):
            sh.add([f'SM{n:04d}', sid[sym], sym, r['asof'], 'FYE1404 (v41)', 'نمای مالی', '', '', '', '', 'فهرست ندارد', '', '', '', '', '', '', '', '',
                    '', 0, 'ندارد', '', '', '', 'NO_ROSTER_AT_ASOF', 'گزارش FY1404 و اطلاعیه ن۴۵ در آرشیو نیست'], fill=FILL_BAD, height=30)
            continue
        nm = norm(r['name']) if not r['name'].startswith('(') else ''
        role = r['role'] or ''
        usable_end = 'بله (مقطع یکسان با پایان مشترک؛ فهرست مبنا نیست)' if m['scd'][sym]['fye'] == END else 'خیر (مقطع متفاوت)'
        sh.add([f'SM{n:04d}', sid[sym], sym, r['asof'], f"FYE1404 (v41)؛ {usable_end}", 'نمای مالی',
                pid(sym, nm) if nm else '', r['name'], nm, f"{(t or {}).get('idbasis') or 'نامشخص'}؛ هش {(t or {}).get('idhash') or ''}",
                'عضو اصلی فعال' if nm else 'نماینده نامشخص؛ خالی بودن واقعی مستند نیست', role,
                'بله' if r['ceo'] == 'بله' else 'خیر', 'بله' if r['chair'] == 'بله' else 'خیر',
                r['holder'] or '', 'ندارد (فقط نام)', r['name'] if r['kind'] and 'نماینده' in r['kind'] else '', norm(r['holder'] or ''),
                'نام دارنده حقوقی' if r['holder'] else 'عضو حقیقی مستقیم؛ بدون تطبیق کرسی',
                '', sum(1 for y in m['roster'] if y['sym'] == sym and not y['id'].endswith('-NA')),
                'نامشخص: تعداد کل اعلام شده در متن ثبت نشده', f"{m['scd'][sym]['src_id']}", m['scd'][sym]['fye'], '',
                '', f"BR Independence v41 ردیف {r['row']}"], height=30)
    sh.autofilter()
    return sh, rr


# person level dates that differ from the seat level date stored in the seat reconstruction
# (the seat row carries the date of the LAST change on the seat, not of this person's own move)
DATE_OVERRIDE = {
    ('تاپیکو', 'عبدالکریم پهلوانی'): ('1404/01/23', 'تاریخ حذف در اطلاعیه 1333972؛ کرسی تا 1404/06/26 بدون نماینده'),
    ('کچاد', 'مجتبی شعبانی'): ('1404/10/29', 'جلسه هیئت مدیره؛ جانشین موقت کریم نژاد از 1404/05/06 مجمع تا 1404/10/29'),
    ('جم', 'امیر اکبری'): ('1404/06/04', 'جابجایی بین کرسی ها؛ فقط سمت و کرسی عوض شد'),
    ('جم', 'محمدهاشم نجفی اردکانی'): ('1404/12/25', 'تاریخ خروج از هیئت؛ در 1404/06/04 فقط بین کرسی ها جابجا شده بود'),
    ('خساپا', 'کورش پرویزیان'): ('1404/02/08', 'جابجایی بین کرسی ها؛ از کرسی رنا به کرسی سازمان مدیریت صنعتی'),
    ('شپدیس', 'سیدعلی نواب فر'): ('', 'نامشخص: جانشین میانی کاردار؛ تاریخ خروج نواب فر در فایل ثبت نشده و پر نشد'),
}


def flows_sheet(wb, m, rr):
    heads = ['شناسه ردیف', 'نماد', 'نمای زمانی', 'شناسه پژوهشی شخص', 'نام خام', 'نام یکسان سازی شده',
             'وضعیت در مبنا', 'وضعیت در پایان', 'کد وضعیت جریان', 'وضعیت جریان', 'سمت در مبنا', 'سمت در پایان',
             'شخص حقوقی مبنا', 'شخص حقوقی پایان', 'جابجایی فرد بین کرسی ها', 'تاریخ موثر یا بازه قابل اثبات تغییر',
             'نوع تاریخ', 'منبع مبنا', 'منبع پایان', 'شاهد تصدی v41 (کلاس و شروع)', 'اطمینان', 'یادداشت']
    sh = Sheet(wb, 'Person Flows v42', 'ورود، خروج و ماندگاری افراد یکتا',
               'عضو جدید یعنی جدید نسبت به مبنا، نه اولین عضویت در تاریخ شرکت. غایب در فهرست ناقص خروج حساب نشد. جابجایی فرد بین دو کرسی ورود و خروج نیست. بخش اول: پنجره مشترک برای هشت شرکت با فهرست اسمی. بخش دوم: نمای مالی FYE1404 با شاهد تصدی v41؛ این بخش فقط می گوید هر عضو پایان دوره نسبت به مبنای قدیمی تازه وارد است یا نه و خروج را نشان نمی دهد.',
               heads, widths={1: 9, 2: 9, 3: 16, 4: 12, 5: 20, 6: 20, 7: 16, 8: 16, 9: 24, 10: 36, 11: 24, 12: 24, 13: 28, 14: 28, 15: 16, 16: 22, 17: 18, 18: 12, 19: 12, 20: 26, 21: 12, 22: 44},
               freeze='D4', raw_cols=(5, 13, 14))
    n = 0
    for sym, (b, e, meta) in m['snaps'].items():
        pf = m['flows'][sym]
        bseat = {s.rep: s for s in b.seats if s.rep}
        eseat = {s.rep: s for s in e.seats if s.rep}
        # raw names and dates from seat reconstruction rows
        rawn, dates, srcb, srce = {}, {}, {}, {}
        for r in m['seat']:
            if r['sym'] != sym:
                continue
            nm = norm(r['rep'] or '')
            cut = S._cutoff(r['snapshot'])
            rawn[nm] = r['rep']
            if cut == BASE:
                srcb[nm] = str(r['tracing'])
                if r['exit'] and nm:
                    dates[nm] = r['exit'].replace('خروج/جایگزینی ', '')
            elif cut == END:
                srce[nm] = str(r['tracing'])
                if r['entry'] and nm:
                    dates[nm] = r['entry'].replace('ورود ', '')
        for name, code in pf['records']:
            n += 1
            bs, es = bseat.get(name), eseat.get(name)
            moved = ''
            if bs and es and (bs.seat_key != es.seat_key):
                moved = 'بله'
            elif bs and es:
                moved = 'خیر'
            raw = rawn.get(name, name)
            date = dates.get(name, '')
            dtype = 'تاریخ موثر ثبت شده در بازسازی v38' if date else ''
            if code == 'PERSISTED' and not (sym, name) in DATE_OVERRIDE:
                date, dtype = '', ''
            if (sym, name) in DATE_OVERRIDE:
                date, dtype = DATE_OVERRIDE[(sym, name)]
                if not date:
                    date, dtype = 'نامشخص', dtype
            if code == 'ABSENT_AT_END_UNDETERMINED':
                date = 'نامشخص؛ بازه v38: پس از 1403/12/29 تا پیش از 1404/08/28'
                dtype = 'بازه؛ تاریخ ساختگی وارد نشد'
            conf = {'FULL': 'قطعی', 'FULL_SENSITIVITY': 'قطعی با حساسیت', 'CONDITIONAL': 'مشروط'}[meta['end']]
            if code in ('ABSENT_AT_END_UNDETERMINED',):
                conf = 'مشروط'
            sh.add([f'PF{n:04d}', sym, 'پنجره مشترک', pid(sym, name), raw, name,
                    'حاضر' if bs else 'غایب', 'حاضر' if es else ('غایب' if meta['end'] != 'CONDITIONAL' else 'غایب در فهرست مشروط'),
                    code, FLOW_FA[code], bs.role if bs else '', es.role if es else '',
                    bs.raw_holder if bs else '', es.raw_holder if es else '', moved, date, dtype,
                    srcb.get(name, ''), srce.get(name, ''), '', conf,
                    meta['caveat'] if (code != 'PERSISTED' and meta['caveat']) else ''],
                   fill=(FILL_WARN if conf != 'قطعی' else None), height=40)
    # fiscal view rows
    for x in rr:
        r = x['r']; t = x['t']; sym = r['sym']
        if r['id'].endswith('-NA'):
            n += 1
            sh.add([f'PF{n:04d}', sym, 'نمای مالی FYE1404', '', '', '', 'نامشخص', 'فهرست ندارد', 'FISCAL_NO_ROSTER', FLOW_FA['FISCAL_NO_ROSTER'],
                    '', '', '', '', '', '', '', '', '', '', 'ندارد', 'گزارش FY1404 و ن۴۵ در آرشیو نیست'], fill=FILL_BAD)
            continue
        if r['name'].startswith('('):
            n += 1
            sh.add([f'PF{n:04d}', sym, 'نمای مالی FYE1404', '', r['name'], '', 'نامشخص', 'کرسی حقوقی با نماینده نامشخص', 'FISCAL_BASE_UNDETERMINED',
                    FLOW_FA['FISCAL_BASE_UNDETERMINED'], '', r['role'], '', r['holder'], '', '', '', '', f"{m['scd'][sym]['src_id']}", 'UNKNOWN', 'ندارد',
                    'نماینده معلوم نیست؛ به خروج یا صفر تبدیل نشد'], fill=FILL_WARN)
            continue
        n += 1
        code = REL2FISCAL[x['rel'][0]]
        nm = norm(r['name'])
        tdesc = f"{t['cls']}؛ شروع {t['start'] or 'نامشخص'}" if t else ''
        if t and t['cls'] in ('INTERVAL',):
            tdesc = f"{t['cls']}؛ بازه {t['lower']} تا {t['upper']}"
        if t and t['cls'] in ('MIN_CONTINUOUS', 'CONFLICT_MIN'):
            tdesc = f"{t['cls']}؛ لنگر {t['upper'] or t['lower']}"
        same_end = m['scd'][sym]['fye'] == END
        sh.add([f'PF{n:04d}', sym, 'نمای مالی FYE1404', pid(sym, nm), r['name'], nm,
                'نامشخص (فهرست اسمی مبنا در دسترس نیست)', f"حاضر در مقطع {r['asof']}", code, FLOW_FA[code], '', r['role'], '', r['holder'] or '', '', '', '', '',
                f"{m['scd'][sym]['src_id']}", tdesc, 'استنباطی از تصدی', x['rel'][1] + ('' if same_end else '؛ فقط تا افق FYE')],
               fill=(FILL_OK if code == 'FISCAL_ENTRY_AFTER_BASE' else None), height=34)
    sh.autofilter()
    return sh


def seatchanges_sheet(wb, m):
    heads = ['شناسه ردیف', 'نماد', 'پنجره زمانی', 'کلید کرسی', 'مبنای تطبیق کرسی', 'دارنده حقوقی در مبنا', 'دارنده حقوقی در پایان',
             'نماینده در مبنا', 'نماینده در پایان', 'سمت در مبنا', 'سمت در پایان', 'کد نوع تغییر', 'نوع تغییر',
             'جابجایی فرد از یا به کرسی دیگر', 'تاریخ موثر یا بازه تغییر', 'منبع مبنا', 'منبع پایان', 'وضعیت کفایت تطبیق', 'یادداشت']
    sh = Sheet(wb, 'Seat Changes v42', 'تغییرات کرسی، شخص حقوقی و نماینده',
               'کرسی فقط با نام دارنده حقوقی یا جایگزینی مستند دارنده جفت شده است؛ شماره ردیف جدول یا ترتیب نام ها مبنا نیست. کرسی حقیقی مستقیم هرگز جفت نشد. تغییر نام دارنده با شناسه ثابت تغییر دارنده نیست؛ شناسه ملی اشخاص حقوقی در فایل نیست و مبنای تطبیق فقط نام است. عبارت فاقد نماینده به خالی بودن، خروج یا صفر تغییر تبدیل نشد.',
               heads, widths={1: 9, 2: 9, 3: 16, 4: 26, 5: 24, 6: 26, 7: 26, 8: 18, 9: 18, 10: 22, 11: 22, 12: 20, 13: 28, 14: 14, 15: 16, 16: 10, 17: 10, 18: 20, 19: 44},
               freeze='D4', raw_cols=(6, 7, 8, 9))
    n = 0
    adit = {}
    for a in m['audit']:
        if a['var'] == 'رویداد کرسی':
            adit[(a['sym'], norm(a['holder'] or ''))] = a
    for sym, (b, e, meta) in m['snaps'].items():
        for r in m['seatrows'][sym]:
            n += 1
            bb, ee = r['base'], r['end']
            a = adit.get((sym, r['key']))
            eff = (a['eff'] if a and a['eff'] and a['eff'] != '—' else '')
            if a is None and '→' in r['key']:
                for kk, aa in adit.items():
                    if kk[0] == sym and '→' in (aa['holder'] or '') and norm(aa['holder']) == r['key']:
                        a = aa; eff = aa['eff'] or ''
            basis = 'نام دارنده حقوقی' if (bb and ee and bb.holder == ee.holder) else 'جایگزینی دارنده طبق ستون قبلی و جدید اطلاعیه (v38)؛ بازبینی نشد'
            moved = 'ورودی از کرسی دیگر' if r['moved_in'] else ''
            if r['moved_out']:
                moved = (moved + ' و ' if moved else '') + 'خروجی به کرسی دیگر'
            adeq = 'کافی' if meta['end'] != 'CONDITIONAL' else 'مشروط'
            if sym == 'پارسان':
                adeq = 'کافی با حساسیت مجمع 1404/10/23'
            if sym == 'جم':
                adeq = 'کافی با حساسیت انتشار پس از پایان'
            if sym == 'کچاد' and 'خوزستان' in r['key']:
                eff = 'جانشین موقت 1404/05/06؛ شعبانی 1404/10/29'
            note = ''
            if r['type'] == 'FILLED_VACANT':
                note = 'کرسی در مبنا خالی مستند بود؛ ورود جدید نسبت به مبنا'
            if r['type'] == 'PERSON_SWAP':
                note = 'جابجایی همان افراد بین کرسی ها'
            sh.add([f'SC{n:04d}', sym, 'پنجره مشترک', r['key'], basis, bb.raw_holder if bb else '', ee.raw_holder if ee else '',
                    bb.raw_rep if bb else '', ee.raw_rep if ee else '', bb.role if bb else '', ee.role if ee else '',
                    r['type'], E.SEAT_TYPES[r['type']], moved, eff, bb.src if bb else '', ee.src if ee else '', adeq, note], height=40)
    # direct seats: Tejarat exit, never paired
    n += 1
    sh.add([f'SC{n:04d}', 'وتجارت', 'پنجره مشترک', 'عضو حقیقی مستقیم', 'ممنوع: بدون شناسه پایدار کرسی', '', '', 'حمیدرضا شریفیان', 'خارج شده؛ کرسی بدون جایگزین',
            'نایب رئیس (موظف)', '', 'UNMATCHED', E.SEAT_TYPES['UNMATCHED'], '', 'نامشخص', '1159936', '1444154', 'ناکافی',
            'کرسی مستقیم؛ خالی شدن مستند کرسی یا حذف کرسی ثابت نشد؛ فقط رویداد فردی در Person Flows v42 ثبت شد'], fill=FILL_WARN, height=40)
    # Hokeshti holder level
    for r, (a_, b_) in zip(m['hk_rows'], [(x, y) for x, y in zip(m['hk_base'].seats, m['hk_end'].seats)]):
        n += 1
        bb, ee = a_, b_
        changed = bb.holder != ee.holder
        sh.add([f'SC{n:04d}', 'حکشتی', 'پنجره مشترک', r['key'], 'جایگزینی دارنده طبق خوانش v38' if changed else 'نام دارنده حقوقی',
                bb.raw_holder, ee.raw_holder, 'فاقد نماینده', 'فاقد نماینده', '', '', 'UNMATCHED',
                'تطبیق نامشخص برای نماینده؛ ' + ('تغییر شخص حقوقی دارنده' if changed else 'دارنده حقوقی ثابت'), '',
                '1404/08/12' if changed else '', bb.src, ee.src, 'ناکافی برای نماینده؛ کافی با محدودیت برای دارنده',
                'نماینده نامشخص؛ به صفر تغییر یا خروج تبدیل نشد'], fill=FILL_WARN, height=40)
    sh.autofilter()
    return sh
