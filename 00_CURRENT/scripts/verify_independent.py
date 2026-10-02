#!/usr/bin/env python3
"""Independent re-computation of the v42 person flows and legacy numerators/denominators.

It reads ONLY the finished workbook (sheet 'Snapshot Members v42' plus the master sheet of v41
that it contains) with plain sets and counters; it does not import the engine in bsi_v42/.

usage: python verify_independent.py <v42.xlsx>
"""
import sys, json
import openpyxl

NAMED = 'عضو اصلی فعال'
VACANT = 'کرسی بدون نماینده؛ خالی بودن مستند'


def main(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Snapshot Members v42']
    snaps = {}
    for r in ws.iter_rows(min_row=4, values_only=True):
        if not r[0] or r[5] != 'پنجره مشترک':
            continue
        sym, cut, status = r[2], r[3], r[10]
        if r[9] and 'نماینده حقیقی نامشخص' in str(r[9]):
            continue                                        # Hokeshti holder level rows
        snaps.setdefault(sym, {}).setdefault(cut, []).append(dict(name=r[8], status=status, holder=r[14], key=r[17], role=r[11]))
    fs = wb['Flow Summary v42']
    fsr = {r[1]: r for r in fs.iter_rows(min_row=4, values_only=True) if r[1]}
    master = {r[0]: r for r in wb['Turnover نهایی'].iter_rows(min_row=4, values_only=True) if r[0]}
    out = dict(companies=len(snaps), flow_diffs=[], legacy_diffs=[], checked_numbers=0)
    for sym, d in snaps.items():
        b = [x for x in d['1403/12/29'] if x['status'] == NAMED]
        e = [x for x in d['1404/12/29'] if x['status'] == NAMED]
        bs, es = {x['name'] for x in b}, {x['name'] for x in e}
        exits, entries, per = len(bs - es), len(es - bs), len(bs & es)
        row = fsr[sym]
        if row[9] not in (None, ''):                        # complete lists: counts present
            for lab, mine, theirs in (('base', len(bs), row[6]), ('end', len(es), row[7]), ('persisted', per, row[8]),
                                       ('exits', exits, row[9]), ('entries', entries, row[10])):
                out['checked_numbers'] += 1
                if mine != theirs:
                    out['flow_diffs'].append((sym, lab, mine, theirs))
            assert len(bs) - exits + entries == len(es)
        # legacy numbers (seat based) from seat keys
        direct = any(not x['key'] for x in b + e if x['holder'] and 'مستقیم' in str(x['holder']))
        mm = master[sym]
        G, H, J, K, M, N, P, Q, S, T = (mm[6], mm[7], mm[9], mm[10], mm[12], mm[13], mm[15], mm[16], mm[18], mm[19])
        if direct:
            mine = dict(G=len(bs - es), H=len(bs), P=len(es - bs), Q=len(es), S=len(bs & es), T=len(bs))
            theirs = dict(G=G, H=H, P=P, Q=Q, S=S, T=T)
        else:
            bk = {x['key']: x for x in d['1403/12/29']}; ek = {x['key']: x for x in d['1404/12/29']}
            both = [k for k in bk if k in ek]
            comp = [k for k in both if bk[k]['status'] == NAMED and ek[k]['status'] == NAMED]
            def holder_of(x, side):
                h = x['holder'] or ''
                return h.split('→')[0].strip() if ('→' in h and side == 'b') else (h.split('→')[1].strip() if '→' in h else h)
            stable = [k for k in comp if holder_of(bk[k], 'b') == holder_of(ek[k], 'e')]
            mine = dict(G=sum(1 for k in comp if bk[k]['name'] != ek[k]['name']), H=len(comp), K=len(stable),
                        J=sum(1 for k in stable if bk[k]['name'] != ek[k]['name']), N=len(both),
                        M=sum(1 for k in both if holder_of(bk[k], 'b') != holder_of(ek[k], 'e')), P=len(es - bs), Q=len(es))
            mine['S'] = mine['H'] - mine['G']; mine['T'] = mine['H']
            theirs = dict(G=G, H=H, J=J, K=K, M=M, N=N, P=P, Q=Q, S=S, T=T)
        for k, v in mine.items():
            out['checked_numbers'] += 1
            if v != theirs[k]:
                out['legacy_diffs'].append((sym, k, v, theirs[k]))
    out['ok'] = not out['flow_diffs'] and not out['legacy_diffs']
    print(json.dumps(out, ensure_ascii=False, indent=1))
    return out


if __name__ == '__main__':
    main(sys.argv[1])
