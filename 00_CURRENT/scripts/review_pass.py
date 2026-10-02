#!/usr/bin/env python3
"""Second, separate review pass over the five sensitive cases and a sample of other companies.

This is an AI re-check of facts quoted in the v42 text against the cells of the workbook they
come from.  It is NOT an independent human evaluation and it does not re-open primary documents.

usage: python review_pass.py <v41.xlsx> <v42.xlsx>
"""
import sys, json
import openpyxl
sys.path.insert(0, __file__.rsplit('/', 1)[0])
from bsi_v42 import load_v41 as L

BASE, END = '1403/12/29', '1404/12/29'


def main(v41, v42):
    wb = L.open_wb(v41); w2 = openpyxl.load_workbook(v42, data_only=True)
    reg = L.read_registry(wb); ten = L.read_tenure(wb); master = L.read_master(wb); scope = {s['sym']: s for s in L.read_scope(wb)}
    notices = lambda sym: sorted((r['pub'][:10], r['tracing']) for r in reg if r['sym'] == sym and r['type'].startswith('اطلاعیه'))
    checks = []
    def chk(case, text, ok, detail=''):
        checks.append(dict(case=case, check=text, ok=bool(ok), detail=str(detail)[:200]))
    # --- Tejarat
    n = notices('وتجارت')
    chk('وتجارت', 'هیچ اطلاعیه ثبت شده بین 1402/12/08 و 1404/09/04 نیست', not [x for x in n if '1402/12/08' < x[0] < '1404/09/04'], n)
    chk('وتجارت', '1443078 انتشار 1404/09/04 و 1444154 انتشار 1404/09/15', dict((t, d) for d, t in n).get(1443078) == '1404/09/04' and dict((t, d) for d, t in n).get(1444154) == '1404/09/15')
    t58 = [t for t in ten if t['id'] == 'T059'][0]
    chk('وتجارت', 'شروع ریاست ترک تبریزی EXACT 1401/12/27', t58['cls'] == 'EXACT' and t58['start'] == '1401/12/27', t58['start'])
    ms = w2['Snapshot Members v42']
    base = [r for r in ms.iter_rows(min_row=4, values_only=True) if r[2] == 'وتجارت' and r[5] == 'پنجره مشترک' and r[3] == BASE and r[10] == 'عضو اصلی فعال']
    end = [r for r in ms.iter_rows(min_row=4, values_only=True) if r[2] == 'وتجارت' and r[5] == 'پنجره مشترک' and r[3] == END and r[10] == 'عضو اصلی فعال']
    chk('وتجارت', 'مبنا پنج نفر و پایان چهار نفر؛ فقط شریفیان غایب', len(base) == 5 and len(end) == 4 and {r[8] for r in base} - {r[8] for r in end} == {'حمیدرضا شریفیان'}, (len(base), len(end)))
    # --- Vpasar
    n = notices('وپاسار')
    chk('وپاسار', 'تنها اطلاعیه ثبت شده 1039576 با انتشار 1402/04/05', n == [('1402/04/05', 1039576)], n)
    chk('وپاسار', 'FYE1404 تایید نشده و فهرست v41 ندارد', scope['وپاسار']['fye_status'] == 'تایید نشده' and any(r['member_id'] == 'F28-NA' for r in ten))
    chk('وپاسار', 'REF05 مستر بدون مقدار', master['وپاسار']['G'] is None)
    # --- Hokeshti
    ro = [r for r in L.read_roster_v41(wb) if r['sym'] == 'حکشتی']
    chk('حکشتی', 'پنج دارنده حقوقی با نماینده نامشخص در v41', len(ro) == 5 and all(r['name'].startswith('(') for r in ro), len(ro))
    chk('حکشتی', 'تفاوت روز مقطع 1404/03/31 تا مبنا برابر 94', scope['حکشتی']['fye'] == '1404/03/31')
    ns = dict((t, d) for d, t in notices('حکشتی'))
    chk('حکشتی', '1310210 انتشار 1403/12/01 و 1433770 انتشار 1404/08/18', ns.get(1310210) == '1403/12/01' and ns.get(1433770) == '1404/08/18')
    # --- Fmli
    chk('فملی', '17 اطلاعیه ثبت شده با نماد فملی', len(notices('فملی')) == 17 and all(r['sym'] == 'فملی' for r in reg if r['sid'].startswith('T19_') and r['tracing'] in {t for _, t in notices('فملی')}))
    get = {t['id']: t for t in ten}
    chk('فملی', 'سه عضو با شروع EXACT پس از مبنا (T001 T003 T006)', all(get[i]['cls'] == 'EXACT' and get[i]['start'] > BASE for i in ('T001', 'T003', 'T006')), [get[i]['start'] for i in ('T001', 'T003', 'T006')])
    chk('فملی', 'خلوصی INTERVAL با بازه شامل مبنا', get['T004']['cls'] == 'INTERVAL' and get['T004']['lower'] <= BASE <= get['T004']['upper'])
    chk('فملی', 'مستر 0 از 5 برای گردش و عضو جدید', (master['فملی']['G'], master['فملی']['H'], master['فملی']['P'], master['فملی']['Q']) == (0, 5, 0, 5))
    chk('فملی', 'شافعی رئیس و فیض اردکانی مدیرعامل و نایب رئیس (دو نقش جدا)', get['T002']['name'] == 'غلامرضا شافعی' and get['T007']['kind'] == 'CEO' and get['T007']['name'] == 'سیدمصطفی فیض اردکانی')
    # --- Rampena
    chk('رمپنا', 'اولی EXACT 1404/02/21 عضو و رئیس', get['T063']['start'] == '1404/02/21' and get['T064']['kind'] == 'CHAIR' and get['T064']['start'] == '1404/02/21')
    nr = notices('رمپنا')
    chk('رمپنا', 'بین 1203308 و 1343229 اطلاعیه ای نیست', not [x for x in nr if '1403/03/26' < x[0] < '1404/02/23'], [x for x in nr if x[0] >= '1403/03/26'][:3])
    mm = master['رمپنا']
    chk('رمپنا', 'مستر 2/5 و 1/1 و 4/5 و 2/5', (mm['G'], mm['H'], mm['J'], mm['K'], mm['M'], mm['N'], mm['P'], mm['Q']) == (2, 5, 1, 1, 4, 5, 2, 5))
    # --- sample of other companies: Flow Summary legacy values equal master and tenure cross-check flag
    fs = {r[1]: r for r in w2['Flow Summary v42'].iter_rows(min_row=4, values_only=True) if r[1]}
    for sym in ('فولاد', 'ومعادن', 'نوری', 'اخابر', 'شتران', 'وامید', 'همراه'):
        mmx = master[sym]
        chk('نمونه ' + sym, 'REF05 قدیمی در Flow Summary برابر مستر و هم راستایی تصدی برقرار', fs[sym][16] == mmx['G'] and fs[sym][17] == mmx['H'] and str(fs[sym][30]).startswith('هم راستا'), (fs[sym][16], fs[sym][17], fs[sym][30]))
    me = w2['Measures v42']
    cnt = {}
    for r in me.iter_rows(min_row=4, values_only=True):
        if r[1] == 'UPX1':
            cnt[r[14]] = cnt.get(r[14], 0) + 1
    chk('پوشش', 'UPX1: کامل 7، مشروط 1، ناکافی 22', cnt == {'کامل': 7, 'مشروط': 1, 'ناکافی': 22}, cnt)
    cr = L.read_cr1a(wb)
    chk('پوشش', 'پوشش محافظه کارانه v39 برابر 26', sum(1 for v in cr.values() if v['conservative'] == 'بله') == 26)
    res = dict(note='بازبینی هوش مصنوعی؛ ارزیابی مستقل انسانی نیست و سند اصلی بازدریافت نشد', n_checks=len(checks), n_failed=sum(1 for c in checks if not c['ok']), checks=checks)
    print(json.dumps(res, ensure_ascii=False, indent=1))
    return res


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
