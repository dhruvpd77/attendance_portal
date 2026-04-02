"""Parse simple paper-setting duty Excel: Date, Subject, Faculty (initial), Notes."""
from __future__ import annotations

import re
from datetime import date, datetime

import openpyxl


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


def parse_paper_setting_workbook(file_obj) -> list[dict]:
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    hdr_idx = None
    col_date = col_subj = col_fac = col_notes = col_deadline = None

    for i, row in enumerate(rows[:40]):
        if not row:
            continue
        lower_cells = [str(c).strip().lower() if c is not None else '' for c in row[:16]]
        for j, lc in enumerate(lower_cells):
            if lc in ('date', 'date of duty', 'assigned date'):
                col_date = j
            if 'subject' in lc and 'total' not in lc:
                col_subj = j
            if lc in ('faculty', 'name of faculty', 'evaluator', 'initial', 'short name', 'faculty initial'):
                col_fac = j
            if 'note' in lc or 'remark' in lc:
                col_notes = j
            if lc in ('deadline', 'due date', 'last date', 'submit by', 'target date'):
                col_deadline = j
        if col_fac is not None and col_subj is not None:
            hdr_idx = i
            break

    if hdr_idx is None:
        hdr_idx = 0
        col_date, col_subj, col_fac, col_notes = 0, 1, 2, 3

    out: list[dict] = []
    for r in range(hdr_idx + 1, len(rows)):
        row = rows[r]
        if not row:
            continue
        row = list(row) + [None] * 16

        fac = row[col_fac] if col_fac is not None and col_fac < len(row) else None
        if fac is None or str(fac).strip() == '':
            continue

        d = None
        if col_date is not None and col_date < len(row):
            d = _cell_date(row[col_date])

        dl = None
        if col_deadline is not None and col_deadline < len(row):
            dl = _cell_date(row[col_deadline])

        subj = ''
        if col_subj is not None and col_subj < len(row) and row[col_subj] is not None:
            subj = str(row[col_subj]).strip()

        notes = ''
        if col_notes is not None and col_notes < len(row) and row[col_notes] is not None:
            notes = str(row[col_notes]).strip()

        out.append(
            {
                'duty_date': d,
                'deadline_date': dl,
                'subject_name': subj,
                'faculty_initial': re.sub(r'\s+', ' ', str(fac).strip()),
                'notes': notes,
            }
        )

    if not out:
        raise ValueError(
            'No rows with a faculty column found. Expected columns like Date, Subject, Faculty (initial), Notes.'
        )
    return out
