"""
Canonical exam phase ordering for UI and Excel: T1, T2, T3, … (by number), then SEE, then others A–Z.
"""
import re

_T_NUM = re.compile(r'^T(\d+)$', re.IGNORECASE)
# T1 in "SY_I - T1", "SY_1_T1", "Lesson T2", etc.
_T_EMBED = re.compile(r'(?:^|[-_\s/—·•])T\s*(\d+)', re.IGNORECASE)
_SEE = re.compile(r'\bSEE\b', re.IGNORECASE)


def exam_phase_header_short_name(name):
    """
    Short Excel/UI band label: T1, T2, …, SEE — strips stream prefixes like "SY_I - T1".
    Used for the compiled 'All departments' sheet so row 1 shows only T1…SEE once each.
    Unknown names are returned unchanged (trimmed).
    """
    raw = (name or '').strip()
    if not raw:
        return ''
    if _SEE.search(raw):
        return 'SEE'
    m = _T_EMBED.search(raw)
    if m:
        return f'T{int(m.group(1))}'
    u = raw.upper().replace(' ', '')
    if u == 'SEE':
        return 'SEE'
    m2 = _T_NUM.match(u)
    if m2:
        return f'T{int(m2.group(1))}'
    return raw


def exam_phase_name_sort_key(name):
    """
    Sort key for phase names: T-prefixed numeric phases first (T1, T2, T10, …),
    then SEE, then any other names alphabetically.
    """
    raw = (name or '').strip()
    if not raw:
        return (3, 0, '')
    u = raw.upper().replace(' ', '')
    m = _T_NUM.match(u)
    if m:
        return (0, int(m.group(1)), raw.lower())
    if u == 'SEE':
        return (1, 0, raw.lower())
    return (2, 0, raw.lower())


def sorted_phase_names(names):
    """Deduplicate and sort phase name strings."""
    return sorted(set(names), key=exam_phase_name_sort_key)


def sort_exam_phases(phase_qs_or_list):
    """Return a new list of ExamPhase instances in display order."""
    lst = list(phase_qs_or_list)
    lst.sort(key=lambda ep: exam_phase_name_sort_key(ep.name))
    return lst
