"""Measures, Coverage, Sources, Dictionary, Changes, QA."""
import re
from . import engine as E, snapshots as S, prep as P, curated as C, jalali as J
from .names import norm
from .xl import Sheet, FX, FILL_WARN, FILL_BAD, FILL_OK
from openpyxl.utils import get_column_letter as CL

BASE, END = S.BASE, S.END
LEG = [('REF05', 'گردش فردی مبتنی بر کرسی', 'G', 'H'),
       ('REF08', 'تغییر نماینده با شخص حقوقی ثابت', 'J', 'K'),
       ('ROB_LEGAL', 'تغییر شخص حقوقی دارنده کرسی', 'M', 'N'),
       ('ROB_NEW', 'سهم اعضای جدید', 'P', 'Q'),
       ('ROB_STABILITY', 'ثبات کرسی با همان شخص', 'S', 'T')]
DEFS = {
 'REF05': 'صورت: کرسی های قابل مقایسه که شخص نامبرده در دو مقطع فرق دارد؛ مخرج: کرسی های قابل مقایسه (مبنای نامبرده). کرسی حقیقی مستقیم در سطح فرد',
 'REF08': 'صورت: کرسی با همان شخص حقوقی که نماینده فرق دارد؛ مخرج: کرسی های با همان شخص حقوقی و نماینده معلوم در هر دو مقطع',
 'ROB_LEGAL': 'صورت: کرسی هایی که شخص حقوقی دارنده فرق دارد؛ مخرج: کرسی های موجود در هر دو مقطع',
 'ROB_NEW': 'صورت: افراد پایان که در مبنا نبودند؛ مخرج: افراد پایان',
 'ROB_STABILITY': 'صورت: کرسی های قابل مقایسه با همان شخص؛ مخرج: کرسی های قابل مقایسه',
}


def rate(n_, d_):
    return (n_ / d_) if (isinstance(n_, (int, float)) and isinstance(d_, (int, float)) and d_ > 0) else ''


def measures_sheet(wb, m):
    heads = ['نماد', 'کد سنجه', 'نام سنجه', 'رده', 'تعریف صورت و مخرج', 'صورت تاریخی v41', 'مخرج تاریخی v41',
             'صورت قابل استفاده', 'مخرج قابل استفاده', 'نرخ قابل استفاده', 'مبنای زمانی', 'تاریخ مبنا', 'تاریخ پایان',
             'واجد شرایط بر پایه تعریف و ساختار', 'کفایت داده', 'وضعیت شواهد', 'مقدار بازتولیدشده با همان تعریف',
             'مقایسه با مقدار تاریخی', 'توضیح', 'سلول مبدا در v41']
    sh = Sheet(wb, 'Measures v42', 'سنجه های قدیمی، بازتولیدشده و جدید با صورت و مخرج',
               'سه چیز جدا نگه داشته شده است: مقدار قدیمی REF05 و تعریف تاریخی آن، مقدار بازتولیدشده با همان تعریف و سنجه های جدید مبتنی بر افراد یکتا با کدهای UPX1 و UPN1 و UPR1 و UPF1. نرخ خروج افراد یکتا بی اعلام جای REF05 نشده است. نبود داده، شرکت را از جمعیت واجد شرایط خارج نکرده است؛ نبود ساختار (کرسی حقوقی) با نبود داده جداست.',
               heads, widths={1: 9, 2: 14, 3: 30, 4: 22, 5: 46, 14: 18, 15: 16, 16: 18, 17: 18, 18: 28, 19: 50, 20: 22}, freeze='C4')
    rows = {}
    master = m['master']
    for s in m['scope']:
        sym = s['sym']; tier = m['tier'][sym]; mm = master[sym]
        leg = m['legacy'].get(sym)
        for code, name, kn, kd in LEG:
            elig = 'بله'
            if code in ('REF08', 'ROB_LEGAL'):
                if sym in P.DIRECT_STRUCT:
                    elig = 'خیر: غیرقابل اعمال'
                elif sym in P.STRUCT_UNDETERMINED:
                    elig = 'نامشخص'
            hn, hd = mm[kn], mm[kd]
            un, ud = '', ''
            suff, status, comp, reprod, why = 'ناکافی', C.NOT_FOUND, '', '', ''
            if elig.startswith('خیر'):
                suff, status, why = 'غیرقابل اعمال', C.NA, 'کرسی ها عضو حقیقی مستقیم اند؛ شخص حقوقی دارنده وجود ندارد'
            elif tier in ('A_FULL', 'A_SENS'):
                un, ud = leg[kn], leg[kd]
                suff, status = 'کامل', C.PARTIAL
                reprod = f'{un} از {ud}'
                comp = 'برابر' if (un, ud) == (hn, hd) else 'تفاوت'
                why = 'بازتولید با تعریف تاریخی از بازسازی سطح کرسی v38؛ سند اصلی در این جلسه بازدریافت نشد'
                if tier == 'A_SENS':
                    why += '؛ حساسیت: ' + m['snaps'][sym][2]['caveat'][:140]
            elif tier == 'A_COND':
                un, ud = leg[kn], leg[kd]
                suff, status = 'مشروط', C.PARTIAL
                reprod = f'{un} از {ud}'
                comp = 'برابر' if (un, ud) == (hn, hd) else 'تفاوت'
                why = 'بازتولید برابر است ولی فهرست پایان وتجارت مشروط است؛ خروج شریفیان قطعی نیست'
            elif tier == 'B_AGG':
                if hn is not None and hd is not None:
                    un, ud = hn, hd
                    suff, status = 'کامل ارثی', C.PARTIAL
                    why = 'ارثی از مستر؛ فهرست اسمی دو مقطع در دسترس نیست و بازتولید اسمی ممکن نشد'
                    comp = 'بازتولید ممکن نشد'
                else:
                    suff, status = 'ناکافی', C.NOT_FOUND
                    why = 'مقدار مستر وجود ندارد'
            elif tier == 'B_CONTRA':
                suff, status = 'نیازمند بررسی', C.CONFLICT
                comp = 'متعارض با شواهد v41'
                why = 'مقدار تاریخی حفظ شد؛ سه عضو با شروع پس از مبنا در v41 با صفر ناسازگار است و فهرست اسمی مبنا برای اصلاح در دسترس نیست'
            elif tier == 'H_HOLDER':
                if code == 'ROB_LEGAL':
                    un, ud = 3, 5
                    suff, status = 'مشروط', C.PARTIAL
                    reprod = '3 از 5 (سطح دارنده)'
                    why = 'سه دارنده تغییر کرده؛ جفت شدن کرسی ها طبق خوانش v38؛ مقدار مستر نداشت'
                else:
                    suff, status = 'ناکافی', C.NOT_FOUND
                    why = 'نماینده ها در اسناد فاقد نماینده اند؛ این کمبود داده است نه غیرقابل اعمال'
            elif tier == 'C_NONE':
                suff, status, why = 'ناکافی', C.NOT_FOUND, 'در منابع بررسی شده فهرست مبنا و پایان قابل اتکا دریافت نشد'
            if code == 'REF08' and sym == 'حکشتی':
                elig = 'بله (دو کرسی حقوقی ثابت)'
            if code == 'ROB_LEGAL' and sym == 'حکشتی':
                elig = 'بله'
            r = sh.r
            sh.add([sym, code, name, 'قدیمی؛ تعریف حفظ شده', DEFS[code], hn if hn is not None else '', hd if hd is not None else '',
                    un, ud, FX(f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r}),I{r}>0),H{r}/I{r},"")', rate(un, ud)),
                    'پنجره مشترک', BASE, END, elig, suff, status, reprod, comp, why,
                    f"Turnover نهایی {mm[kn + '_cell']} و {mm[kd + '_cell']}"],
                   fill=(FILL_BAD if suff in ('نیازمند بررسی',) else (FILL_WARN if suff in ('مشروط', 'ناکافی') else None)), height=48)
            rows[(sym, code)] = r
        # CR1A
        v = mm['chair_change']
        raw = 1 if v == 'بله' else (0 if v == 'خیر' else '')
        cr = m['cr1a'][sym]
        if tier in ('A_FULL', 'A_SENS', 'A_COND'):
            ch = m['chairs'][sym]
            un = ch[0]; suff, status = ('کامل', C.PARTIAL) if tier != 'A_COND' else ('مشروط', C.PARTIAL)
            reprod = f"رئیس مبنا {ch[1]}؛ رئیس پایان {ch[2]}"
            comp = 'برابر' if un == raw else 'تفاوت'
            why = 'مقایسه هویت رئیس در دو مقطع از بازسازی سطح کرسی v38؛ تعداد رویدادهای میانی جدا در Turnover Events v42'
        elif sym == 'فملی':
            un = ''; suff, status = 'نیازمند بررسی', C.CONFLICT
            reprod = 'رئیس پایان: غلامرضا شافعی (از 1404/02/02)؛ رئیس مبنا نامعلوم'
            comp = 'متعارض'
            why = 'مقدار خام خیر حفظ شد؛ شواهد غیرمستقیم v41 بله است؛ در پوشش محافظه کارانه v39 نبود و نماند'
        elif sym == 'رمپنا':
            un = ''; suff, status = 'پشتیبانی غیرمستقیم', C.PARTIAL
            reprod = 'رئیس پایان: محمدرضا اولی (از 1404/02/21)؛ رئیس مبنا نامعلوم'
            comp = 'مستر نامشخص بود'
            why = 'عضویت و ریاست رئیس پایان پس از مبنا آغاز شده؛ تغییر بله غیرمستقیم؛ هویت رئیس مبنا نیازمند اطلاعیه 1203308 و 1343229'
        elif sym in ('وپاسار', 'حکشتی'):
            un = ''; suff, status = 'ناکافی', C.NOT_FOUND
            why = 'هویت رئیس در دو مقطع معلوم نیست'
        else:
            un = raw if cr['conservative'] == 'بله' else ''
            suff, status = 'کامل ارثی', C.PARTIAL
            why = 'ارثی از CR1A v39 (' + str(cr['status']) + ')؛ تایید مجدد مستقل همه منابع انجام نشد'
            comp = 'بازتولید ممکن نشد'
        r = sh.r
        sh.add([sym, 'CR1A', 'تغییر رئیس هیئت مدیره بین دو مقطع', 'قدیمی؛ تعریف حفظ شده',
                'صورت: 1 اگر هویت رئیس در دو مقطع فرق دارد؛ مخرج: 1. جدا از تعداد رویدادهای میانی', raw, 1 if raw != '' else '',
                un, 1 if un != '' else '', FX(f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r}),I{r}>0),H{r}/I{r},"")', rate(un, 1 if un != '' else '')),
                'پنجره مشترک', BASE, END, 'بله', suff, status, reprod, comp if tier in ('A_FULL', 'A_SENS', 'A_COND') else comp if sym in ('فملی', 'رمپنا') else comp if tier == 'B_AGG' else '', why,
                f"Turnover نهایی V{mm['row']}؛ CR1A v39 ردیف {cr['row']}"],
               fill=(FILL_BAD if suff == 'نیازمند بررسی' else (FILL_WARN if suff in ('مشروط', 'ناکافی', 'پشتیبانی غیرمستقیم') else None)), height=48)
        rows[(sym, 'CR1A')] = r
        # new unique person measures
        pf = m['flows'].get(sym)
        for code, name, key, den in (('UPX1', 'نرخ خروج افراد یکتا', 'exits', 'n_base'),
                                      ('UPN1', 'سهم افراد واردشده', 'entries', 'n_end'),
                                      ('UPR1', 'نرخ ماندگاری افراد یکتا', 'persisted', 'n_base')):
            r = sh.r
            if pf and pf['complete']:
                un, ud = pf[key], pf[den]
                suff, status = 'کامل', C.PARTIAL
                why = 'هر دو فهرست اسمی کامل؛ هویت ها با نام یکسان سازی شده در همان شرکت مطابقت داده شد'
                if m['tier'][sym] == 'A_SENS':
                    why += '؛ حساسیت ثبت شده در Five Cases و Snapshot Members'
                reprod = f'{un} از {ud}'
            elif pf:
                un, ud = '', ''
                suff, status = 'مشروط', C.PARTIAL
                reprod = ''
                why = 'فهرست پایان مشروط است؛ نسبت دقیق کل هیئت ساخته نشد. اگر خروج شریفیان تایید شود خروج 1 از 5 و ماندگاری 4 از 5 است'
            elif sym == 'حکشتی':
                un, ud, reprod = '', '', ''
                suff, status, why = 'ناکافی', C.NOT_FOUND, 'نماینده ها نامشخص؛ بدون فهرست اسمی نرخی ساخته نشد'
            elif sym == 'وپاسار':
                un, ud, reprod = '', '', ''
                suff, status, why = 'ناکافی', C.NOT_FOUND, 'در منابع بررسی شده فهرست قابل اتکا دریافت نشد'
            else:
                un, ud, reprod = '', '', ''
                suff, status = 'ناکافی', C.UNAVAILABLE
                why = 'فهرست اسمی مبنا در بسته شواهد v41 است که در این جلسه ارائه نشد؛ شمار تجمیعی مستر برای افراد یکتا کافی نیست'
            sh.add([sym, code, name, 'جدید؛ کد غیرمتداخل', {'UPX1': 'صورت: افراد حاضر در مبنا و غایب در پایان؛ مخرج: اعضای فعال مبنا',
                                                          'UPN1': 'صورت: افراد حاضر در پایان و غایب در مبنا؛ مخرج: اعضای فعال پایان',
                                                          'UPR1': 'صورت: افراد حاضر در هر دو مقطع؛ مخرج: اعضای فعال مبنا'}[code],
                    '', '', un, ud, FX(f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r}),I{r}>0),H{r}/I{r},"")', rate(un, ud)),
                    'پنجره مشترک', BASE, END, 'بله', suff, status, reprod, '', why, 'Snapshot Members v42؛ Person Flows v42'],
                   fill=(FILL_WARN if suff in ('مشروط', 'ناکافی') else None), height=44)
            rows[(sym, code)] = r
        # fiscal view
        r = sh.r
        sh.add([sym, 'UPF1', 'افراد یکتا در نمای مالی (FY1403 تا FY1404)', 'جدید؛ نمای مالی',
                'همان سه نسبت UPX1 و UPN1 و UPR1 ولی بین پایان FY1403 و FY1404 هر شرکت؛ فقط با دو فهرست کامل معتبر', '', '', '', '',
                FX(f'=IF(AND(ISNUMBER(H{r}),ISNUMBER(I{r}),I{r}>0),H{r}/I{r},"")', ''),
                'نمای مالی', m['tb'][sym]['fy3'], m['scd'][sym]['fye'], 'بله', 'ناکافی', 'هنوز بررسی نشده', '', '',
                'فهرست اسمی پایان FY1403 در فایل نیست؛ نتیجه نمای مالی نامشخص ماند و عدد پنجره مشترک جای آن نیامد', 'Scope v42'], height=44)
        rows[(sym, 'UPF1')] = r
    sh.autofilter()
    return sh, rows


def coverage_sheet(wb, m, mrows):
    heads = ['معیار', 'کد', 'مبنای زمانی', 'کل شرکت ها', 'واجد شرایط بر پایه تعریف و ساختار', 'دارای داده کافی کامل',
             'کامل ارثی', 'مشروط', 'پشتیبانی غیرمستقیم', 'نیازمند بررسی', 'ناکافی', 'غیرقابل اعمال', 'یادداشت']
    sh = Sheet(wb, 'Coverage v42', 'پوشش به تفکیک تعریف، مبنای زمانی و کفایت داده',
               'برای هر سنجه سه شمارش آمده است: کل 30 شرکت، واجد شرایط بر پایه تعریف و ساختار، و دارای داده کافی. شمارش ها با فرمول از Measures v42 گرفته شده اند. پنجره مشترک و نمای مالی در هم جمع نشده اند. کامل ارثی یعنی عدد مستر بدون بازتولید اسمی.',
               heads, widths={1: 34, 2: 12, 3: 18, 13: 60}, freeze='D4')
    ms = "'Measures v42'"
    res = {}
    names = {'REF05': 'REF05 گردش فردی مبتنی بر کرسی', 'REF08': 'REF08 تغییر نماینده با شخص حقوقی ثابت', 'ROB_LEGAL': 'ROB_LEGAL تغییر شخص حقوقی',
             'ROB_NEW': 'ROB_NEW سهم اعضای جدید', 'ROB_STABILITY': 'ROB_STABILITY ثبات', 'CR1A': 'CR1A تغییر رئیس (پوشش ارثی و بازتولیدشده)',
             'UPX1': 'UPX1 نرخ خروج افراد یکتا', 'UPN1': 'UPN1 سهم افراد واردشده', 'UPR1': 'UPR1 نرخ ماندگاری افراد یکتا', 'UPF1': 'UPF1 نمای مالی افراد یکتا'}
    # compute python counts from the same rows by re-reading cells of ws
    ws = wb['Measures v42']
    def val(r, c):
        return ws.cell(r, c).value
    notes = {
        'REF08': 'سه شرکت با عضو حقیقی مستقیم غیرقابل اعمال؛ وپاسار نامشخص؛ حکشتی واجد شرایط است و کمبود داده دارد (مستر آن را اشتباه خیر نوشته)',
        'ROB_LEGAL': 'حکشتی در سطح دارنده 3 از 5 مشروط است',
        'CR1A': 'پوشش محافظه کارانه v39 دست نخورد: 26 از 30. فملی و رمپنا شواهد غیرمستقیم دارند و وارد آن نشدند',
        'UPX1': 'فقط هشت شرکت فهرست اسمی دو مقطع دارند؛ یک شرکت مشروط است', 'UPF1': 'نمای مالی برای هیچ شرکتی محاسبه نشد چون فهرست FY1403 در فایل نیست',
        'REF05': 'فملی نیازمند بررسی؛ وپاسار و حکشتی بدون مقدار؛ وتجارت مشروط'}
    for code in names:
        r = sh.r
        cnt = lambda col, crit: f'COUNTIFS({ms}!$B$4:$B$2000,"{code}",{ms}!${col}$4:${col}$2000,"{crit}")'
        # python values
        lst = [rr for (sym, c), rr in mrows.items() if c == code]
        tot = len(lst)
        py = lambda col, crit: sum(1 for rr in lst if str(val(rr, col)) == crit)
        elig_py = sum(1 for rr in lst if str(val(rr, 14)).startswith('بله'))
        vals = [names[code], code, 'نمای مالی' if code == 'UPF1' else 'پنجره مشترک',
                FX(f'=COUNTIF({ms}!$B$4:$B$2000,"{code}")', tot),
                FX(f'=COUNTIFS({ms}!$B$4:$B$2000,"{code}",{ms}!$N$4:$N$2000,"بله*")', elig_py),
                FX('=' + cnt('O', 'کامل'), py(15, 'کامل')), FX('=' + cnt('O', 'کامل ارثی'), py(15, 'کامل ارثی')),
                FX('=' + cnt('O', 'مشروط'), py(15, 'مشروط')), FX('=' + cnt('O', 'پشتیبانی غیرمستقیم'), py(15, 'پشتیبانی غیرمستقیم')),
                FX('=' + cnt('O', 'نیازمند بررسی'), py(15, 'نیازمند بررسی')), FX('=' + cnt('O', 'ناکافی'), py(15, 'ناکافی')),
                FX('=' + cnt('O', 'غیرقابل اعمال'), py(15, 'غیرقابل اعمال')), notes.get(code, '')]
        sh.add(vals, height=40)
        res[code] = dict(total=tot, eligible=elig_py, full=py(15, 'کامل'), inherited=py(15, 'کامل ارثی'), cond=py(15, 'مشروط'),
                         indirect=py(15, 'پشتیبانی غیرمستقیم'), review=py(15, 'نیازمند بررسی'), insuff=py(15, 'ناکافی'), na=py(15, 'غیرقابل اعمال'))
    sh.add([''] * 13)
    # extra rows: two lists complete by time basis
    tiers = m['tier']
    sh.add(['شرکت های دارای دو فهرست اسمی کامل: پنجره مشترک، کل', 'LISTS2', 'پنجره مشترک', 30,
            FX(f'=COUNTIFS({ms}!$B$4:$B$2000,"UPX1",{ms}!$O$4:$O$2000,"کامل")', res['UPX1']['full']),
            res['UPX1']['full'], '', res['UPX1']['cond'], '', '', '', '', 'پنج شرکت بدون حساسیت و دو شرکت (جم و پارسان) با حساسیت ثبت شده؛ وتجارت مشروط'], height=40)
    for rel in sorted({v['rel'] for v in m['tb'].values()}):
        syms = [s for s in m['syms'] if m['tb'][s]['rel'] == rel]
        full = [s for s in syms if tiers[s] in ('A_FULL', 'A_SENS')]
        cond = [s for s in syms if tiers[s] == 'A_COND']
        sh.add([f'به تفکیک رابطه مبنای زمانی: {rel}', 'LISTS2_BY_BASIS', 'پنجره مشترک', len(syms), '', len(full), '', len(cond), '', '', '', '', 'شرکت های کامل: ' + '، '.join(full)], height=40)
    sh.add(['شرکت های دارای دو فهرست اسمی کامل: نمای مالی FY1403 تا FY1404', 'LISTS2_FISCAL', 'نمای مالی', 30, '', 0, '', 0, '', '', 30, '', 'فهرست اسمی FY1403 در فایل نیست'], height=40)
    sh.add(['CR1A پوشش محافظه کارانه (تعریف v39)', 'CR1A_CONS', 'پنجره مشترک', 30, 30, 26, '', '', 2, '', 2, '', 'بله فملی و رمپنا غیرمستقیم؛ وپاسار و حکشتی بدون کد'], height=40)
    sh.autofilter()
    return sh, res
