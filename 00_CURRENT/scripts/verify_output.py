#!/usr/bin/env python3
"""Re-open the built v42 file and check it against v41 (read only) and against its own cached values.

usage: python verify_output.py <v41.xlsx> <v42.xlsx>
"""
import sys, openpyxl, json


def main(v41, v42):
    a_f = openpyxl.load_workbook(v41); a_v = openpyxl.load_workbook(v41, data_only=True)
    b_f = openpyxl.load_workbook(v42); b_v = openpyxl.load_workbook(v42, data_only=True)
    res = dict(sheets_v41=len(a_f.sheetnames), sheets_v42=len(b_f.sheetnames))
    diff_formula = diff_cached = 0; ncell = 0
    examples = []
    for ws in a_f:
        w2 = b_f[ws.title]
        for row in ws.iter_rows():
            for c in row:
                ncell += 1
                v2 = w2[c.coordinate].value
                v1 = c.value
                if hasattr(v1, 'text'):          # ArrayFormula
                    v1 = (v1.ref, v1.text); v2 = (getattr(v2, 'ref', None), getattr(v2, 'text', None))
                if v1 != v2:
                    diff_formula += 1
                    if len(examples) < 5:
                        examples.append((ws.title, c.coordinate, str(v1)[:40], str(v2)[:40]))
        wv1 = a_v[ws.title]; wv2 = b_v[ws.title]
        for row in wv1.iter_rows():
            for c in row:
                x, y = c.value, wv2[c.coordinate].value
                if x != y and not (isinstance(x, (int, float)) and isinstance(y, (int, float)) and abs(x - y) < 1e-9):
                    diff_cached += 1
                    if len(examples) < 8:
                        examples.append(('cached', ws.title, c.coordinate, str(x)[:30], str(y)[:30]))
    res.update(old_cells_compared=ncell, old_formula_or_value_diffs=diff_formula, old_cached_value_diffs=diff_cached, examples=examples)
    # new sheets: formula cells must have cached values, no error strings
    new = [n for n in b_f.sheetnames if n not in a_f.sheetnames]
    nform = nmissing = nerr = 0
    for n in new:
        for row in b_f[n].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    nform += 1
                    v = b_v[n][c.coordinate].value
                    if v is None:
                        nmissing += 1          # empty-string results read back as None
                    if isinstance(v, str) and v.startswith('#'):
                        nerr += 1
    res.update(new_sheets=new, new_formulas=nform, new_formula_cached_empty=nmissing, new_error_values=nerr)
    errs = 0
    for ws in b_v:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('#') and ws.title in new:
                    errs += 1
    res['error_cells_in_new_sheets'] = errs
    res['rtl'] = {n: b_f[n].sheet_view.rightToLeft for n in new}
    print(json.dumps(res, ensure_ascii=False, indent=1))
    return res


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
