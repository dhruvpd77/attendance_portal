"""
Parse LJIET-style combined supervision Excel (faculty × subject columns with dates/times).
"""
from __future__ import annotations

import re
from datetime import date, datetime

import openpyxl

from core.models import Department, Faculty, InstituteSemester
from core.exam_subunit_scope import canonical_division_code

# Auto-created per academic semester when supervision/paper sheets name unknown faculty.
VISITING_DEPARTMENT_NAME = 'Visiting faculty'


def _norm(s) -> str:
    if s is None:
        return ''
    return re.sub(r'\s+', ' ', str(s).strip()).upper()


def match_faculty_for_department(department, full_name: str, short_initial: str) -> Faculty | None:
    """Match faculty by full name or short initial within department."""
    fn = _norm(full_name)
    si = _norm(short_initial)
    qs = list(Faculty.objects.filter(department=department))
    if fn:
        for f in qs:
            if _norm(f.full_name) == fn:
                return f
        fn_compact = fn.replace(' ', '')
        for f in qs:
            if _norm(f.full_name).replace(' ', '') == fn_compact:
                return f
    if si:
        for f in qs:
            if _norm(f.short_name) == si:
                return f
    return None


def match_faculty_global(full_name: str, short_initial: str) -> Faculty | None:
    """First match by full name or short initial across all departments (hub-phase uploads)."""
    fn = _norm(full_name)
    si = _norm(short_initial)
    qs = list(Faculty.objects.select_related('department').all().order_by('department_id', 'pk'))
    if fn:
        for f in qs:
            if _norm(f.full_name) == fn:
                return f
        fn_compact = fn.replace(' ', '')
        for f in qs:
            if _norm(f.full_name).replace(' ', '') == fn_compact:
                return f
    if si:
        for f in qs:
            if _norm(f.short_name) == si:
                return f
    return None


def ensure_visiting_faculty(institute_semester: InstituteSemester, full_name: str, short_initial: str) -> Faculty:
    """Create or reuse a Faculty row under auto-created Visiting faculty department for this academic semester."""
    dept, _ = Department.objects.get_or_create(
        institute_semester=institute_semester,
        name=VISITING_DEPARTMENT_NAME,
        defaults={},
    )
    fac = match_faculty_for_department(dept, full_name, short_initial)
    if fac:
        return fac

    fn = (full_name or '').strip()
    si = (short_initial or '').strip()
    if not fn and not si:
        fn = 'Unknown visitor'
        si = 'VIS'
    elif not fn:
        fn = si
    elif not si:
        si = re.sub(r'[^A-Z0-9]', '', _norm(fn))[:8] or 'VST'
    si = si[:30]
    if not si:
        si = 'VST'

    if Faculty.objects.filter(department=dept, full_name__iexact=fn).exists():
        return Faculty.objects.filter(department=dept, full_name__iexact=fn).first()

    base = si
    n = 0
    while Faculty.objects.filter(department=dept, short_name__iexact=si).exists():
        n += 1
        si = f'{base[:25]}-{n}'[:30]

    return Faculty.objects.create(
        department=dept,
        full_name=fn[:200],
        short_name=si[:30],
    )


def resolve_supervision_faculty_for_phase(
    *,
    institute_semester: InstituteSemester,
    coordinator_department: Department | None,
    hub_phase: bool,
    full_name: str,
    short_initial: str,
) -> tuple[Faculty, bool]:
    """
    Match sheet name to Faculty; if not found in any regular workflow, attach to Visiting faculty department.
    Returns (faculty, created_or_matched_via_visiting).
    """
    if hub_phase:
        fac = match_faculty_global(full_name, short_initial)
    else:
        fac = None
        if coordinator_department:
            fac = match_faculty_for_department(coordinator_department, full_name, short_initial)
        if not fac:
            fac = match_faculty_global(full_name, short_initial)
    if fac:
        return fac, False
    return ensure_visiting_faculty(institute_semester, full_name, short_initial), True


def resolve_faculty_with_visiting_fallback(
    institute_semester: InstituteSemester | None,
    *,
    coordinator_department: Department | None,
    scope_wide: bool,
    full_name: str,
    short_initial: str,
) -> Faculty | None:
    """Paper duties / shared resolver: match globally or locally, then Visiting faculty if semester known."""
    if scope_wide:
        fac = match_faculty_global(full_name, short_initial)
    else:
        fac = None
        if coordinator_department:
            fac = match_faculty_for_department(coordinator_department, full_name, short_initial)
        if not fac:
            fac = match_faculty_global(full_name, short_initial)
    if fac:
        return fac
    if not institute_semester:
        return None
    return ensure_visiting_faculty(institute_semester, full_name, short_initial)


def _cell_date(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()[:10]
    try:
        return datetime.strptime(s, '%Y-%m-%d').date()
    except ValueError:
        return None


def parse_combined_supervision_workbook(file_obj) -> list[dict]:
    """
    Returns list of dicts:
      faculty_name, faculty_initial, subject_name, supervision_date, time_slot, division_code
    """
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    hdr_idx = None
    for i, row in enumerate(rows):
        if not row:
            continue
        for cell in row:
            if cell and 'Name of Faculty' in str(cell):
                hdr_idx = i
                break
        if hdr_idx is not None:
            break

    if hdr_idx is None or hdr_idx < 3:
        raise ValueError(
            'Could not find a header row containing "Name of Faculty". '
            'Use the standard combined supervision sheet layout.'
        )

    subject_row = rows[hdr_idx - 3]
    date_row = rows[hdr_idx - 2]
    time_row = rows[hdr_idx - 1]

    duty_cols = []
    max_c = max(len(subject_row), len(date_row), len(time_row))
    for c in range(max_c):
        subj = subject_row[c] if c < len(subject_row) else None
        if subj is None or str(subj).strip() == '':
            continue
        u = str(subj).strip().upper()
        if u in ('SUB', 'TOTAL'):
            continue
        duty_cols.append(c)

    if not duty_cols:
        raise ValueError('No subject columns found above the date row.')

    assignments: list[dict] = []

    for r in range(hdr_idx + 1, len(rows)):
        row = rows[r]
        if not row or len(row) < 2:
            continue
        name = row[1]
        if name is None or str(name).strip() == '':
            continue
        nup = str(name).strip().upper()
        if 'TOTAL' in nup and len(nup) < 20:
            continue

        initial = row[2] if len(row) > 2 else ''

        for c in duty_cols:
            if c >= len(row):
                continue
            cell_val = row[c]
            if cell_val is None or str(cell_val).strip() == '':
                continue

            dt = _cell_date(date_row[c] if c < len(date_row) else None)
            if not dt:
                continue

            subj = subject_row[c] if c < len(subject_row) else None
            tm = time_row[c] if c < len(time_row) else None
            subj_name = str(subj).strip() if subj else ''
            time_s = str(tm).strip() if tm else ''

            division = canonical_division_code(cell_val)

            assignments.append({
                'faculty_name': str(name).strip(),
                'faculty_initial': str(initial).strip() if initial else '',
                'subject_name': subj_name,
                'supervision_date': dt,
                'time_slot': time_s,
                'division_code': division,
            })

    if not assignments:
        raise ValueError('No supervision assignments found in the sheet (check dates and duty cells).')

    return assignments
