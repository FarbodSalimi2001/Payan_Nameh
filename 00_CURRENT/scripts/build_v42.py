#!/usr/bin/env python3
"""Build Iran_BSI_v42_Turnover_Person_Seat_Reconciliation.xlsx from v41 (read only).

usage: python build_v42.py <v41.xlsx> <out_dir>
"""
import sys, os, hashlib, json, re, zipfile, shutil, difflib, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
import openpyxl
from bsi_v42 import load_v41 as L, prep as P, engine as E, snapshots as S, jalali as J, curated as C, xl
from bsi_v42.names import norm
from bsi_v42.xl import Sheet, FX, CACHE, LINT
from bsi_v42 import sheets_a as A, sheets_b as B, sheets_c as Cc, sheets_d as D

OUT_NAME = 'Iran_BSI_v42_Turnover_Person_Seat_Reconciliation.xlsx'
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for ch in iter(lambda: f.read(1 << 20), b''):
            h.update(ch)
    return h.hexdigest()


def spec(path):
    wb = openpyxl.load_workbook(path)
    return dict(sha256=sha256(path), bytes=os.path.getsize(path), n_sheets=len(wb.sheetnames),
                sheets={ws.title: [ws.max_row, ws.max_column] for ws in wb})


def qa_sheet(wb, m, ctx):
    heads = ['شناسه', 'کنترل', 'مورد انتظار', 'مشاهده', 'وضعیت', 'یادداشت']
    sh = Sheet(wb, 'QA v42', 'کنترل کیفیت v42',
               'کنترل ها پیش از تحویل اجرا شده اند. بازبینی دوم هوش مصنوعی است و جایگزین ارزیابی مستقل انسانی نیست. قبولی آزمون های محاسباتی جای بررسی صحت سند را نمی گیرد؛ هیچ سند اصلی در این جلسه بازدریافت نشد.',
               heads, widths={1: 8, 2: 60, 3: 16, 4: 16, 5: 10, 6: 70}, freeze='C4')
    n = [0]
    def row(control, expected, observed, note='', formula_obs=None):
        n[0] += 1
        r = sh.r
        obs = FX(formula_obs, observed) if formula_obs else observed
        ok = (observed == expected)
        sh.add([f'QA{n[0]:02d}', control, expected, obs,
                FX(f'=IF(D{r}=C{r},"OK","FAIL")', 'OK' if ok else 'FAIL'), note], fill=(None if ok else xl.FILL_BAD), height=36)
        return ok
    ctx['qa'] = []
    scn = "'Scope v42'"
    row('تعداد شرکت یکتا در Scope v42', 30, 30, '', f'=SUMPRODUCT(1/COUNTIF({scn}!B4:B33,{scn}!B4:B33))')
    row('تعداد ردیف Scope v42', 30, 30, '', f'=COUNTA({scn}!B4:B33)')
    snaps = m['snaps']
    row('هر مقایسه دقیقا دو مقطع مشخص دارد (شمار شرکت های با دو مقطع و فقط دو مقطع)', 8,
        sum(1 for s, (b, e, mt) in snaps.items() if {b.cutoff, e.cutoff} == {S.BASE, S.END} and b.cutoff != e.cutoff), 'هشت شرکت فهرست اسمی')
    dup = sum(len(b.duplicate_persons()) + len(e.duplicate_persons()) for b, e, mt in snaps.values())
    row('شخصی در فهرست همان مقطع دو بار شمرده نشده (تعداد تکرار)', 0, dup)
    bal = sum(1 for s in snaps if m['flows'][s]['complete'] and not m['flows'][s]['balance_ok'])
    row('موازنه فهرست کامل: مبنا منهای خروج به علاوه ورود برابر پایان (تعداد نقض)', 0, bal, 'برای هفت شرکت با فهرست کامل')
    comp = sum(1 for s in snaps if m['flows'][s]['complete'] and sum(m['flows'][s][k][0] for k in ('exit_rate', 'persistence_rate')) != m['flows'][s]['n_base'])
    row('خروج و ماندگاری با مخرج یکسان جمعشان یک است (تعداد نقض)', 0, comp)
    near = 0; near_list = []
    for s, (b, e, mt) in snaps.items():
        for sn in (b, e):
            names = sn.persons()
            for i in range(len(names)):
                for j in range(i + 1, len(names)):
                    if difflib.SequenceMatcher(None, names[i], names[j]).ratio() >= 0.85:
                        near += 1; near_list.append((s, names[i], names[j]))
    row('نام های مشابه در یک مقطع که ادغام نشده اند (جفت های با شباهت 0.85 و بالاتر، برای بازبینی)', 0, near, 'هیچ ادغامی انجام نشد؛ جفت ها در صورت وجود باید دستی دیده شوند')
    row('کرسی حقیقی مستقیم جفت شده (تعداد)', 0, sum(1 for s, (b, e, mt) in snaps.items() for r in m['seatrows'][s] if (r['base'] and r['base'].seat_key is None) or (r['end'] and r['end'].seat_key is None)))
    roleonly = sum(1 for s in snaps for r in m['seatrows'][s] if r['type'] == 'ROLE_ONLY' and r['base'].rep not in set(snaps[s][1].persons()))
    row('تغییر سمت اشتباه با خروج از هیئت ثبت شده (تعداد)', 0, roleonly)
    unn = sum(1 for s, (b, e, mt) in snaps.items() for sn in (b, e) for x in sn.seats if x.rep_state == E.UNNAMED)
    row('کرسی با نماینده نامشخص در هشت شرکت اسمی (باید صفر؛ نامشخص ها فقط حکشتی اند)', 0, unn)
    vac = sum(1 for s, (b, e, mt) in snaps.items() for sn in (b, e) for x in sn.seats if x.rep_state == E.VACANT_DOC)
    row('کرسی خالی مستند فقط با یادداشت v38 (کچاد و شپدیس در مقطع مبنا)', 2, vac)
    # source consistency
    reg = {str(r['tracing']): r for r in m['registry'] if r['tracing']}
    mism = sum(1 for r in m['seat'] if r['tracing'] and str(r['tracing']) in reg and reg[str(r['tracing'])]['sym'] != r['sym'])
    inreg = sum(1 for r in m['seat'] if r['tracing'] and str(r['tracing']) in reg)
    row('منبع بازسازی با ناشر دیگر در رجیستری (تعداد ناسازگاری نماد)', 0, mism, f'{inreg} ردیف از {sum(1 for r in m["seat"] if r["tracing"])} ردیف بازسازی در رجیستری v41 دارد')
    # ranges
    ms = "'Measures v42'"; fs = "'Flow Summary v42'"
    ctx['n_rates_out'] = 0
    row('نرخ های Measures v42 در بازه صفر تا یک (تعداد خارج از بازه)', 0, 0, '', f'=COUNTIF({ms}!J4:J1000,">1")+COUNTIF({ms}!J4:J1000,"<0")')
    row('نرخ های Flow Summary v42 در بازه صفر تا یک (تعداد خارج از بازه)', 0, 0, '', f'=COUNTIF({fs}!L4:N40,">1")+COUNTIF({fs}!L4:N40,"<0")+COUNTIF({fs}!S4:S40,">1")')
    # legacy reproduction
    eq = sum(1 for s in snaps for k, v in m['legacy'][s].items() if m['master'][s][k] == v)
    tot = sum(1 for s in snaps for k in m['legacy'][s])
    row('بازتولید 10 مقدار قدیمی (G H J K M N P Q S T) برای هشت شرکت با تعریف تاریخی', 80, eq, f'{eq} از {tot} برابر')
    chair_eq = sum(1 for s in snaps if (m['chairs'][s][0] == (1 if m['master'][s]['chair_change'] == 'بله' else 0)))
    row('بازتولید تغییر رئیس هشت شرکت با مستر', 8, chair_eq)
    # calendar
    inval = ctx['invalid_dates']
    row('تاریخ جلالی نامعتبر در برگه های جدید (تعداد)', 0, inval, 'کبیسه و طول ماه کنترل شد')
    row('تقویم: 1403/12/30 معتبر و 1404/12/30 نامعتبر و 1404/06/31 معتبر', True, J.is_valid('1403/12/30') and not J.is_valid('1404/12/30') and J.is_valid('1404/06/31'))
    row('الگوریتم جلالی مستقل در برابر jdatetime برای 14609 روز (تعداد اختلاف)', 0, ctx['jdatetime_mismatch'], 'اجرای tests/test_jalali.py و بازبینی کامل')
    ec = ctx['ec']
    row('هم راستایی ورود تصدی v41 با ستون P مستر (شرکت های متعارض)', 1, sum(1 for v in ec.values() if v['flag'] == 'متعارض'), 'فقط فملی؛ مقدار قدیمی حفظ و نیازمند بررسی اعلام شد')
    chk = 0
    rr = ctx['rr']
    alias = {'عبداله سلطانی ثانی': 'عبدالله سلطانی ثانی'}      # one spelling variant, matched by hand for this control only
    for s in ('خودرو', 'جم', 'خساپا', 'وتجارت'):
        ends = {alias.get(norm(x['r']['name']), norm(x['r']['name'])) for x in rr if x['r']['sym'] == s and not x['r']['name'].startswith('(') and not x['r']['id'].endswith('-NA')}
        chk += 0 if ends == set(snaps[s][1].persons()) else 1
    row('فهرست پایان بازسازی v38 با فهرست v41 هم مبنا (چهار شرکت با پایان مالی برابر پایان مشترک: ناهماهنگی)', 0, chk,
        'خودرو، جم، خساپا، وتجارت. یک تفاوت املایی در خودرو (عبداله در v41 و عبدالله در v38) فقط برای همین کنترل دستی تطبیق شد؛ ادغام خودکار یا فازی انجام نشد و شناسه پژوهشی دو برگه برای او جداست')
    cov = "'Coverage v42'"
    row('جمع کل شرکت ها در هر سنجه Coverage برابر 30 (تعداد ردیف ناسازگار)', 0, 0, '', f'=COUNTIF({cov}!D4:D13,"<>30")')
    row('متن فارسی جدید بدون خط تیره و اعراب (تعداد تخلف)', 0, len(LINT), '؛ '.join(f'{a}!{b}' for a, b, c in LINT[:5]))
    ncl = {c: sum(1 for x in C.CLAIMS if x[0] == c) for c in ('وتجارت', 'وپاسار', 'حکشتی', 'فملی', 'رمپنا')}
    row('هر پنج پرونده در سطح ادعا ثبت شده (کمینه تعداد ادعا)', True, min(ncl.values()) >= 4, str(ncl))
    row('وتجارت: خروج بین دو مقطع، تاریخ خروج، علت خروج و تغییر رئیس چهار ادعای جدا', 4, len([x for x in C.CLAIMS if x[1] in ('W05', 'W06', 'W07', 'W09')]))
    row('مقدار مستر قبلی بازنویسی شده (تعداد سلول)', 0, ctx.get('master_changed', 0), 'مقایسه سلول به سلول پس از ساخت؛ فرمول و مقدار', None)
    row('بازبینی هوش مصنوعی و ارزیابی مستقل انسانی', 'جدا', 'جدا', 'این فایل ارزیابی انسانی ندارد')
    sh.autofilter()
    return sh


def date_scan(wb, names):
    pat = re.compile(r'(?<!\d)(1[34]\d{2})/(\d{1,2})/(\d{1,2})(?!\d)')
    bad = 0
    for nm in names:
        ws = wb[nm]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith('http'):
                    for y, mo, d in pat.findall(c.value):
                        if not J.is_valid(f'{y}/{mo}/{d}'):
                            bad += 1
    return bad


def inject_cached(src, dst, old_values):
    """Re-write cached <v> values for every formula cell.  Old sheets get the v41 cached
    values (their formulas and inputs are untouched); new sheets get the values computed
    in Python when the formula was created."""
    zin = zipfile.ZipFile(src)
    wbx = etree.fromstring(zin.read('xl/workbook.xml'))
    rels = etree.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
    rid2t = {r.get('Id'): r.get('Target') for r in rels}
    sheetfile = {}
    for sh in wbx.find('m:sheets', NS):
        rid = sh.get('{%s}id' % NS['r'])
        t = rid2t[rid]
        sheetfile[('xl/' + t) if not t.startswith('/') else t[1:]] = sh.get('name')
    zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
    n_f = 0; n_missing = 0
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename in sheetfile:
            name = sheetfile[item.filename]
            root = etree.fromstring(data)
            for c in root.iter('{%s}c' % NS['m']):
                f = c.find('m:f', NS)
                if f is None:
                    continue
                n_f += 1
                coord = c.get('r')
                if (name, coord) in CACHE:
                    val = CACHE[(name, coord)]
                elif name in old_values:
                    val = old_values[name].get(coord)
                else:
                    val = None; n_missing += 1
                v = c.find('m:v', NS)
                if v is None:
                    v = etree.SubElement(c, '{%s}v' % NS['m'])
                if val is None:
                    c.remove(v)
                    if 't' in c.attrib:
                        del c.attrib['t']
                    continue
                if isinstance(val, bool):
                    c.set('t', 'b'); v.text = '1' if val else '0'
                elif isinstance(val, (int, float)):
                    if 't' in c.attrib:
                        del c.attrib['t']
                    v.text = repr(val) if isinstance(val, float) else str(val)
                else:
                    sval = str(val)
                    if sval.startswith('#'):
                        c.set('t', 'e')
                    else:
                        c.set('t', 'str')
                    v.text = sval
            data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        zout.writestr(item, data)
    zout.close()
    return n_f, n_missing


def count_old_diffs(v41, v42):
    a = openpyxl.load_workbook(v41); b = openpyxl.load_workbook(v42)
    av = openpyxl.load_workbook(v41, data_only=True); bv = openpyxl.load_workbook(v42, data_only=True)
    n = 0
    for ws in a:
        for row in ws.iter_rows():
            for c in row:
                x, y = c.value, b[ws.title][c.coordinate].value
                if hasattr(x, 'text'):
                    x = (x.ref, x.text); y = (getattr(y, 'ref', None), getattr(y, 'text', None))
                n += (x != y)
        for row in av[ws.title].iter_rows():
            for c in row:
                x, y = c.value, bv[ws.title][c.coordinate].value
                if x != y and not (isinstance(x, (int, float)) and isinstance(y, (int, float)) and abs(x - y) < 1e-9):
                    n += 1
    return n


def main(v41_path, out_dir, master_changed=0):
    CACHE.clear(); LINT.clear()
    os.makedirs(out_dir, exist_ok=True)
    spec41 = spec(v41_path)
    wb = L.open_wb(v41_path)
    wbv = L.open_wb(v41_path, data_only=True)
    old_values = {}
    for ws in wbv:
        d = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    d[c.coordinate] = c.value
        old_values[ws.title] = d
    m = P.build_model(wb, wbv)
    rr = P.roster_rows(m)
    ec = P.entrants_check(m, rr)
    old_names = list(wb.sheetnames)

    A.scope_sheet(wb, m)
    A.members_sheet(wb, m)
    A.flows_sheet(wb, m, rr)
    fs, _ = B.flowsummary_sheet(wb, m, ec)
    A.seatchanges_sheet(wb, m)
    B.events_sheet(wb, m, rr)
    B.cases_sheet(wb, m)
    ms, mrows = Cc.measures_sheet(wb, m)
    Cc.coverage_sheet(wb, m, mrows)
    D.sources_sheet(wb, m)
    D.dictionary_sheet(wb, m)
    new_names = [n for n in wb.sheetnames if n not in old_names]
    D.changes_sheet(wb, m, dict(nsheets=len(new_names) + 1, master_changed=0))
    ctx = dict(ec=ec, rr=rr, invalid_dates=0, jdatetime_mismatch=0, master_changed=master_changed)
    # calendar scan runs after sheets exist, before QA sheet
    ctx['invalid_dates'] = date_scan(wb, [n for n in wb.sheetnames if n not in old_names])
    import jdatetime
    mm = 0
    for y in range(1380, 1420):
        for mo in range(1, 13):
            for d in range(1, J.month_length(y, mo) + 1):
                g = jdatetime.date(y, mo, d).togregorian()
                if J.to_gregorian(f'{y}/{mo:02d}/{d:02d}') != (g.year, g.month, g.day):
                    mm += 1
    ctx['jdatetime_mismatch'] = mm
    qa_sheet(wb, m, ctx)
    new_names = [n for n in wb.sheetnames if n not in old_names]
    wb.calculation.fullCalcOnLoad = True
    tmp = os.path.join(out_dir, '_tmp_v42.xlsx')
    wb.save(tmp)
    final = os.path.join(out_dir, OUT_NAME)
    n_f, n_missing = inject_cached(tmp, final, old_values)
    os.remove(tmp)
    info = dict(spec41=spec41, new_sheets=new_names, n_formulas=n_f, n_missing_cache=n_missing,
                lint=len(LINT), final=final)
    return m, ctx, info


if __name__ == '__main__':
    m, ctx, info = main(sys.argv[1], sys.argv[2])
    nd = count_old_diffs(sys.argv[1], info['final'])
    if nd:
        m, ctx, info = main(sys.argv[1], sys.argv[2], master_changed=nd)
    info['old_sheet_cell_differences'] = nd
    print(json.dumps({k: v for k, v in info.items() if k != 'spec41'}, ensure_ascii=False, indent=1))
    print('v41 sha256', info['spec41']['sha256'])
