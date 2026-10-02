"""Persian name normalisation used for same-company identity matching."""
import re

_DIGITS = str.maketrans('۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩', '01234567890123456789')
_MAP = {'ي': 'ی', 'ى': 'ی', 'ك': 'ک', 'ة': 'ه', 'ۀ': 'ه', 'أ': 'ا', 'إ': 'ا', 'ؤ': 'و',
        '‌': ' ', '‏': ' ', '‎': ' ', ' ': ' ', 'ـ': ''}
_DIAC = re.compile('[ً-ٰٟ]')


def norm(s):
    """Normalise a Persian name: unify ye/kaf, ZWNJ->space, strip diacritics,
    collapse whitespace.  Deliberately does NOT drop words, so two different
    names never collapse into one (no fuzzy merging)."""
    if s is None:
        return ''
    t = str(s).translate(_DIGITS)
    for a, b in _MAP.items():
        t = t.replace(a, b)
    t = _DIAC.sub('', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def pid(company, name):
    """Research person id: stable within a company, derived from the normalised
    name (no national id is available in the v41 workbook for these rows)."""
    import hashlib
    h = hashlib.sha1(f'{norm(company)}|{norm(name)}'.encode('utf-8')).hexdigest()[:8].upper()
    return f'P-{h}'
