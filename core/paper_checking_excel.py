"""Parse evaluation / paper-checking duty Excel (SY1–SY4 block columns)."""
from __future__ import annotations

import re
from datetime import date, datetime, timedelta

import openpyxl

from core.models import Department


def _cell_date(val) -> date | None:
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


def resolve_department_from_sheet_code(code: str) -> Department | None:
    if code is None:
        return None
    raw = str(code).strip()
    if not raw:
        return None
    c = raw.upper().replace(' ', '')
    d = Department.objects.filter(name__iexact=raw.strip()).first()
    if d:
        return d
    m = re.match(r'^(SY|FY|TY|CE|IT|ME|EE|EC)(\d+)$', c, re.I)
    if m:
        candidate = f'{m.group(1).upper()}_{m.group(2)}'
        d = Department.objects.filter(name__iexact=candidate).first()
        if d:
            return d
    if re.match(r'^[A-Z]+\d+$', c):
        with_underscore = re.sub(r'(\D+)(\d+)$', r'\1_\2', c)
        d = Department.objects.filter(name__iexact=with_underscore).first()
        if d:
            return d
    return Department.objects.filter(name__icontains=raw.strip()).first()


def parse_paper_checking_workbook(file_obj) -> list[dict]:
    """
    Returns rows:
      exam_date, subject_name, allocations: [{dept_code, block_range}], total_students, evaluator_initial
    """
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    hdr_idx = None
    for i, row in enumerate(rows):
        if not row or len(row) < 2:
            continue
        a0 = str(row[0]).strip() if row[0] is not None else ''
        a1 = str(row[1]).strip() if row[1] is not None else ''
        if a0 == 'Date of Exam' and 'Subject' in a1:
            hdr_idx = i
            break

    if hdr_idx is None:
        raise ValueError(
            'Could not find header row with "Date of Exam" and "Subject". '
            'Use the standard evaluation-duty Excel layout.'
        )

    dept_row = list(rows[hdr_idx + 1]) if hdr_idx + 1 < len(rows) else []
    dept_headers: list[str] = []
    for col_idx in range(2, 6):
        if col_idx < len(dept_row) and dept_row[col_idx] is not None:
            dept_headers.append(str(dept_row[col_idx]).strip())
        else:
            dept_headers.append(f'DEPT{col_idx - 1}')
    while len(dept_headers) < 4:
        dept_headers.append(f'DEPT{len(dept_headers)}')

    current_date: date | None = None
    current_subject = ''
    out: list[dict] = []

    for r in range(hdr_idx + 3, len(rows)):
        row = rows[r]
        if not row:
            continue
        row = list(row) + [None] * max(0, 8 - len(row))

        if row[0] is not None and str(row[0]).strip():
            d = _cell_date(row[0])
            if d:
                current_date = d
        if row[1] is not None and str(row[1]).strip():
            current_subject = str(row[1]).strip()

        allocations = []
        for ci in range(4):
            col_idx = 2 + ci
            cell = row[col_idx] if col_idx < len(row) else None
            if cell is None or str(cell).strip() == '':
                continue
            code = dept_headers[ci] if ci < len(dept_headers) else f'DEPT{ci}'
            allocations.append({'dept_code': code, 'block_range': str(cell).strip()})

        eval_cell = row[7] if len(row) > 7 else None
        if eval_cell is None or str(eval_cell).strip() == '':
            continue

        total_cell = row[6] if len(row) > 6 else None
        try:
            total_students = int(total_cell) if total_cell is not None else 0
        except (TypeError, ValueError):
            total_students = 0

        if current_date is None:
            continue

        out.append(
            {
                'exam_date': current_date,
                'subject_name': current_subject,
                'allocations': allocations,
                'total_students': total_students,
                'evaluator_initial': str(eval_cell).strip(),
            }
        )

    if not out:
        raise ValueError('No data rows found after the header in the Excel file.')
    return out


def default_checking_deadline(exam_date: date) -> date:
    return exam_date + timedelta(days=1)
