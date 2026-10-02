import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from bsi_v42 import jalali as J


def test_leap_and_month_lengths():
    assert J.is_leap(1403) and not J.is_leap(1404)
    assert J.is_valid('1403/12/30') and not J.is_valid('1404/12/30')
    assert J.is_valid('1404/06/31') and not J.is_valid('1404/07/31')
    assert J.last_day_of_year(1403) == '1403/12/30' and J.last_day_of_year(1404) == '1404/12/29'


def test_days_between_common_cutoffs():
    assert J.days_between('1403/12/29', '1404/12/29') == 366
    assert J.add_days('1403/12/29', 1) == '1403/12/30'


def test_against_jdatetime_if_installed():
    try:
        import jdatetime
    except ImportError:
        return
    for y in range(1395, 1410):
        for m in range(1, 13):
            for d in range(1, J.month_length(y, m) + 1):
                g = jdatetime.date(y, m, d).togregorian()
                assert J.to_gregorian(f'{y}/{m:02d}/{d:02d}') == (g.year, g.month, g.day)
