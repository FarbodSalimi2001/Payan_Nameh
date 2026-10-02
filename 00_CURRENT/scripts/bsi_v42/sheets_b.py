"""Flow Summary, Turnover Events, Five Cases, Measures."""
from . import engine as E, snapshots as S, prep as P, curated as C, jalali as J
from .names import norm
from .xl import Sheet, FX, FILL_WARN, FILL_BAD, FILL_OK
from openpyxl.utils import get_column_letter as CL

BASE, END = S.BASE, S.END


def rate(num, den):
    return (num / den) if (isinstance(num, (int, float)) and isinstance(den, (int, float)) and den > 0) else ''


def flowsummary_sheet(wb, m, ec):
    heads = ['شناسه', 'نماد', 'سطح داده', 'مبنای زمانی نتیجه', 'فهرست مبنا', 'فهرست پایان', 'اعضای مبنا (افراد یکتا)',
             'اعضای پایان (افراد یکتا)', 'افراد مشترک', 'افراد خارج شده', 'افراد واردشده',
             'نرخ خروج افراد یکتا (UPX1)', 'سهم افراد واردشده (UPN1)', 'نرخ ماندگاری افراد یکتا (UPR1)',
             'کنترل موازنه مبنا منهای خروج به علاوه ورود برابر پایان', 'کنترل جمع خروج و ماندگاری برابر یک',
             'REF05 قدیمی صورت', 'REF05 قدیمی مخرج', 'نرخ REF05 قدیمی', 'تفاوت با نرخ خروج افراد یکتا و دلیل',
             'کرسی های مشترک در دو مقطع', 'کرسی حقوقی با دارنده ثابت', 'ثابت با نماینده معلوم در هر دو مقطع',
             'ثابت با نماینده نامشخص', 'ثابت با خالی بودن مستند', 'تغییر نماینده REF08 صورت', 'REF08 مخرج',
             'شاهد تصدی v41: ورود پس از مبنا', 'شاهد تصدی v41: نامشخص', 'مستر: عضو جدید P', 'هم راستایی تصدی با مستر', 'یادداشت']
    sh = Sheet(wb, 'Flow Summary v42', 'خلاصه ورود و خروج افراد یکتا به تفکیک شرکت',
               'نرخ ها فقط وقتی هر دو فهرست کامل و بدون نماینده نامشخص باشند ساخته می شوند؛ در غیر این صورت سلول خالی می ماند و صفر نیست. نرخ خروج و نرخ ماندگاری با مخرج یکسان مکمل اند و دو امتیاز مستقل نیستند. جمع ورود و خروج به عنوان نرخ گردش واحد معرفی نشده است. تفاوت با REF05 از جابجایی همان افراد بین کرسی ها می آید.',
               heads, widths={1: 7, 2: 9, 3: 28, 4: 20, 5: 14, 6: 18, 20: 40, 31: 18, 32: 44}, freeze='C4')
    cmap = {'FULL': 'کامل', 'FULL_SENSITIVITY': 'کامل با حساسیت', 'CONDITIONAL': 'مشروط'}
    out = {}
    for s in m['scope']:
        sym = s['sym']; tier = m['tier'][sym]
        r = sh.r
        g = lambda c: f'{CL(c)}{r}'
        vals = [s['id'], sym, P.TIER_LABEL[tier], 'پنجره مشترک']
        num_cells = [None] * 28
        e = ec[sym]
        if sym in m['snaps']:
            b, en, meta = m['snaps'][sym]
            pf = m['flows'][sym]; leg = m['legacy'][sym]; st = m['structure'][sym]
            complete = pf['complete']
            nb, ne = pf['n_base'], pf['n_end']
            ex, enn, per = pf['exits'], pf['entries'], pf['persisted']
            vals += [cmap[meta['base']], cmap[meta['end']], nb, ne, per if complete else '', ex if complete else '', enn if complete else '']
            rx = rate(ex, nb) if complete else ''
            rn = rate(enn, ne) if complete else ''
            rp = rate(per, nb) if complete else ''
            vals += [FX(f'=IF(AND(ISNUMBER({g(10)}),ISNUMBER({g(7)}),{g(7)}>0),{g(10)}/{g(7)},"")', rx),
                     FX(f'=IF(AND(ISNUMBER({g(11)}),ISNUMBER({g(8)}),{g(8)}>0),{g(11)}/{g(8)},"")', rn),
                     FX(f'=IF(AND(ISNUMBER({g(9)}),ISNUMBER({g(7)}),{g(7)}>0),{g(9)}/{g(7)},"")', rp),
                     FX(f'=IF(AND(ISNUMBER({g(7)}),ISNUMBER({g(10)}),ISNUMBER({g(11)}),ISNUMBER({g(8)})),IF({g(7)}-{g(10)}+{g(11)}={g(8)},"OK","FAIL"),"")', 'OK' if complete else ''),
                     FX(f'=IF(AND(ISNUMBER({g(12)}),ISNUMBER({g(14)})),IF(ABS({g(12)}+{g(14)}-1)<0.000001,"OK","FAIL"),"")', 'OK' if complete else '')]
            vals += [leg['G'], leg['H'], FX(f'=IF(AND(ISNUMBER({g(17)}),ISNUMBER({g(18)}),{g(18)}>0),{g(17)}/{g(18)},"")', rate(leg['G'], leg['H']))]
            if complete and ex != leg['G']:
                why = f'REF05 {leg["G"]} از {leg["H"]} تغییر شخص در کرسی را می شمارد و خروج یکتا {ex} است؛ فرق از جابجایی همان افراد بین کرسی ها می آید'
            elif sym == 'وتجارت':
                why = 'کرسی مستقیم؛ REF05 همان خروج فردی است ولی فهرست پایان مشروط است'
            else:
                why = 'برابر'
            vals += [why]
            both = st['seats_both'] if not sym == 'وتجارت' else ''
            vals += [both, st['stable_holder_seats'] if both != '' else '', st['stable_known_both'] if both != '' else '',
                     st['stable_unknown_rep'] if both != '' else '', st['stable_vacant_documented'] if both != '' else '',
                     st['ref08_changed'] if both != '' else '', st['stable_known_both'] if both != '' else '']
        elif sym == 'حکشتی':
            vals += ['کامل برای دارنده ها؛ نماینده ها نامشخص', 'کامل برای دارنده ها؛ نماینده ها نامشخص', '', '', '', '', '']
            vals += ['', '', '', '', '']
            vals += ['', '', '', 'داده کافی پیدا نشد']
            vals += [5, 2, 0, 2, 0, '', 0]
        else:
            leg_ok = m['master'][sym]['G'] is not None
            vals += ['منبع در دسترس نبود', 'منبع در دسترس نبود', '', '', '', '', '']
            vals += ['', '', '', '', '']
            mm = m['master'][sym]
            vals += [mm['G'] if leg_ok else '', mm['H'] if leg_ok else '',
                     FX(f'=IF(AND(ISNUMBER({g(17)}),ISNUMBER({g(18)}),{g(18)}>0),{g(17)}/{g(18)},"")', rate(mm['G'], mm['H']) if leg_ok else '')]
            vals += ['فهرست اسمی مبنا در دسترس نیست؛ نرخ خروج یکتا ساخته نشد' if leg_ok else 'داده کافی پیدا نشد']
            if leg_ok and mm['N'] is not None:
                stable = mm['N'] - mm['M']
                vals += [mm['N'], stable, mm['K'], stable - mm['K'], '', mm['J'], mm['K']]
            else:
                vals += ['', '', '', '', '', '', '']
        vals += [e['entrants'] if e['n'] else '', e['undetermined'] if e['n'] else '', e['P'] if e['P'] is not None else '', e['flag']]
        note = ''
        if sym == 'فملی':
            note = 'مستر 0 از 5 با سه عضو دارای شروع پس از مبنا در v41 متعارض است؛ مقدار قدیمی حفظ و مقدار قابل استفاده نیازمند بررسی'
        if sym == 'جم':
            note = 'حساسیت انتشار: با فهرست منتشرشده تا پایان، ورود 0 و رئیس بدون تغییر می شد'
        if sym == 'پارسان':
            note = 'حساسیت مجمع 1404/10/23: دو دارنده حقوقی تازه و نمایندگان بعد از پایان نامیده شدند'
        if sym == 'وتجارت':
            note = 'فهرست پایان مشروط؛ خروج شریفیان قطعی اعلام نشد'
        vals += [note]
        sh.add(vals, fill=(FILL_WARN if tier in ('A_COND', 'B_CONTRA', 'H_HOLDER') else (FILL_BAD if tier == 'C_NONE' else None)), height=48)
        out[sym] = r
    sh.autofilter()
    return sh, out


def events_sheet(wb, m, rr):
    heads = ['شناسه رویداد', 'نماد', 'نوع رویداد', 'شخص', 'شخص حقوقی یا کرسی', 'از', 'به', 'تاریخ موثر', 'نوع تاریخ',
             'تاریخ انتشار', 'داخل پنجره مشترک', 'در مقایسه دو مقطع دیده می شود', 'افق مشاهده', 'زنجیره رویدادها کامل',
             'منبع', 'اطمینان', 'یادداشت']
    sh = Sheet(wb, 'Turnover Events v42', 'رویدادهای مستند داخل دوره',
               'مقایسه دو فهرست ابتدا و انتها همه رویدادهای میانی را نشان نمی دهد؛ فردی که خارج و دوباره وارد شده در هر دو مقطع حاضر است ولی دو رویداد دارد. تاریخ موثر، تاریخ جلسه و تاریخ انتشار جدا ثبت شده اند و تاریخ نامشخص با روز ساختگی پر نشده است. تا کامل شدن زنجیره، تعداد رویدادهای شناسایی شده تعداد قطعی همه رویدادها نیست. CR1A بر پایه تفاوت هویت رئیس در دو مقطع است و تعداد رویدادهای تغییر رئیس جداست.',
               heads, widths={1: 9, 2: 9, 3: 22, 4: 24, 5: 28, 6: 18, 7: 18, 8: 13, 9: 22, 10: 13, 11: 12, 12: 14, 13: 20, 14: 16, 15: 18, 16: 12, 17: 50},
               freeze='D4')
    n = 0
    in_window = lambda d: ('بله' if (d and BASE < d <= END) else ('خارج از پنجره' if d else 'نامشخص'))
    for ev in C.EVENTS:
        (sym, typ, who, seat, frm, to, eff, dtype, pub, seen, src, conf, note) = ev
        n += 1
        sh.add([f'EV{n:04d}', sym, typ, who, seat, frm, to, eff, dtype, pub, in_window(eff), seen,
                'تا 1404/12/29', 'تحلیل v38؛ بازبینی سند در این جلسه ممکن نشد', src, conf, note], height=45)
    # tenure derived events for companies without named two-cutoff lists
    for x in rr:
        r, t = x['r'], x['t']; sym = r['sym']
        if sym in m['snaps'] or t is None:
            continue
        horizon = m['scd'][sym]['fye']
        if t['cls'] == 'EXACT' and BASE < t['start'] <= END:
            n += 1
            sh.add([f'EV{n:04d}', sym, 'شروع عضویت (ورود نسبت به مبنا)', r['name'], r['holder'] or '', '', r['name'], t['start'],
                    {'BOARD_MEETING_DATE_IN_NOTICE': 'تاریخ جلسه در اطلاعیه'}.get(t['datetype'], t['datetype'] or ''),
                    '', 'بله', 'خیر یا نامشخص (فهرست مبنا نیست)', f'تا {horizon}', 'ناقص: فقط تصدی اعضای حاضر در پایان دوره',
                    f"BR Tenure v41 {t['id']}", 'متوسط', f"{t['basis']}؛ خروج های متناظر در این جدول نیست"], height=40)
        elif t['cls'] == 'INTERVAL' and t['lower'] and BASE < t['lower']:
            n += 1
            sh.add([f'EV{n:04d}', sym, 'شروع عضویت (ورود نسبت به مبنا)', r['name'], r['holder'] or '', '', r['name'], '',
                    'بازه؛ تاریخ دقیق نامشخص', '', 'بله', 'نامشخص', f'تا {horizon}', 'ناقص',
                    f"BR Tenure v41 {t['id']}", 'پایین', f"بازه {t['lower']} تا {t['upper']}؛ تاریخ ساختگی وارد نشد"], height=40)
    for t in m['tenure']:
        sym = t['sym']
        if t['kind'] != 'CHAIR' or sym in m['snaps']:
            continue
        if t['cls'] == 'EXACT' and BASE < t['start'] <= END:
            n += 1
            sh.add([f'EV{n:04d}', sym, 'شروع ریاست (رئیس جدید نسبت به مبنا)', t['name'], t['holder'] or '', '', t['name'], t['start'],
                    'تاریخ جلسه در اطلاعیه', '', 'بله', 'بله' if m['scd'][sym]['fye'] == END else 'نامشخص', f"تا {m['scd'][sym]['fye']}",
                    'ناقص', f"BR Tenure v41 {t['id']}", 'متوسط', 'نام رئیس مبنا در فایل نیست؛ فقط شروع ریاست فعلی پس از مبنا ثبت شد'], height=40)
    sh.autofilter()
    return sh, n


def cases_sheet(wb, m):
    heads = ['شناسه ادعا', 'پرونده', 'ادعا', 'مقطع یا زمان', 'وضعیت نهایی ادعا', 'یافته v42', 'منابع در فایل', 'سند یا کار لازم',
             'اثر بر سنجه ها', 'یادداشت']
    sh = Sheet(wb, 'Five Cases v42', 'نتیجه پنج پرونده تاریخی در سطح هر ادعا',
               'هر پرونده به ادعاهای جدا شکسته شد: خروج تاییدشده بین دو مقطع، تاریخ دقیق خروج، علت خروج و تغییر رئیس چهار ادعای جدا با وضعیت مستقل اند. پرونده ای که فقط بخشی از آن حل شده کامل حل شده اعلام نشده است. هیچ سند اصلی در این جلسه بازدریافت نشد؛ بسته شواهد v41 ارائه نشد و کدال با 503 و قطع اتصال پاسخ داد، بنابراین وضعیت ها از متن و بازسازی داخل خود فایل اند و تایید مستقل متن سند نیستند.',
               heads, widths={1: 8, 2: 9, 3: 40, 4: 22, 5: 16, 6: 70, 7: 34, 8: 34, 9: 28, 10: 30}, freeze='D4')
    for case, st, txt in C.CASE_SUMMARY:
        sh.add(['SUM', case, 'خلاصه پرونده', '', st, txt, '', '', '', ''], fill=FILL_WARN, height=60)
    for row in C.CLAIMS:
        case, cid, claim, scope, status, finding, srcs, need, eff, note = row
        fill = FILL_BAD if status in (C.CONFLICT, C.UNAVAILABLE, C.NOT_FOUND) else None
        sh.add([cid, case, claim, scope, status, finding, srcs, need, eff, note], fill=fill, height=95)
    sh.autofilter()
    return sh
