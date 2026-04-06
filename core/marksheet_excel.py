"""Detect columns and import exam marksheets (.xlsx) — legacy headers and Subject (Phase) Out of … format."""
from __future__ import annotations

import re
from typing import Any

from core.models import ExamPhase, ExamPhaseSubject, Student, StudentMark, Subject
from core.student_marks_utils import normalize_student_mark

BRANCH_MAX_LEN = 80


def normalize_enrollment_cell(val: Any) -> str | None:
    if val is None:
        return None
    if isinstance(val, float):
        if val != val:  # NaN
            return None
        if val == int(val):
            val = int(val)
    s = str(val).strip()
    if len(s) > 2 and s.endswith('.0') and s[:-2].replace('-', '').isdigit():
        s = s[:-2]
    s = ''.join(s.split())
    if not s.isdigit():
        return None
    return s


def _cell_str(c: Any) -> str:
    if c is None:
        return ''
    return str(c).strip()


def _norm_header_for_match(h: str) -> str:
    """Collapse whitespace/newlines so e.g. 'TOC (T1)\\nOut of 25' matches the compiled pattern."""
    return re.sub(r'\s+', ' ', (h or '').strip())


def detect_marksheet_columns(
    ws,
    subject_name: str,
    phase_name: str,
    *,
    max_header_row: int = 30,
) -> tuple[int | None, int | None, int | None, int]:
    """
    Find header row in the first ``max_header_row`` rows.

    Returns (enroll_col, branch_col or None, marks_col, data_start_row).
    If no header row matched, returns fallbacks enroll_col=3, marks_col=6, data_start_row=1.
    """
    enroll_col: int | None = None
    branch_col: int | None = None
    marks_col: int | None = None
    data_start_row = 0

    subj = (subject_name or '').strip()
    phase = (phase_name or '').strip()
    dyn_pattern: re.Pattern[str] | None = None
    if subj and phase:
        dyn_pattern = re.compile(
            rf'^{re.escape(subj)}\s*\(\s*{re.escape(phase)}\s*\)\s*out\s*of\b',
            re.IGNORECASE,
        )

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        if row_idx > max_header_row:
            break
        if not row:
            continue
        row_lower = [str(c).lower().strip() if c is not None else '' for c in row]
        joined = ' '.join(row_lower)
        if 'enroll' not in joined and 'enrolllment' not in joined:
            continue

        enroll_col = None
        branch_col = None
        marks_col = None

        for c_idx, cell_orig in enumerate(row):
            cell_lo = row_lower[c_idx] if c_idx < len(row_lower) else ''
            hdr = _cell_str(cell_orig)
            if 'enroll' in cell_lo or 'enrolllment' in cell_lo:
                enroll_col = c_idx
                continue
            if 'branch' in cell_lo and 'enroll' not in cell_lo:
                branch_col = c_idx

        for c_idx, cell_orig in enumerate(row):
            hdr = _norm_header_for_match(_cell_str(cell_orig))
            if dyn_pattern and hdr and dyn_pattern.search(hdr):
                marks_col = c_idx
                break
        if marks_col is None:
            for c_idx, cell_orig in enumerate(row):
                cell_lo = row_lower[c_idx] if c_idx < len(row_lower) else ''
                if 'mark' in cell_lo and 'enroll' not in cell_lo:
                    marks_col = c_idx
                    break

        if enroll_col is not None and marks_col is not None:
            data_start_row = row_idx
            break

    if enroll_col is None:
        enroll_col = 3
    if marks_col is None:
        marks_col = 6
    if data_start_row == 0:
        data_start_row = 1

    return enroll_col, branch_col, marks_col, data_start_row


def process_exam_marksheet_worksheet(
    ws,
    *,
    template_phase: ExamPhase,
    template_subject: Subject,
    department_ids: list[int],
) -> dict[str, int]:
    """
    Read active worksheet: match students by enrollment_no across ``department_ids``;
    for each row resolve that student's department-local phase/subject by same names
    as template_phase / template_subject. Optionally set ``Student.branch`` from a BRANCH
    column when the field is empty.
    """
    subj_name = template_subject.name
    phase_name = template_phase.name

    enroll_col, branch_col, marks_col, data_start_row = detect_marksheet_columns(
        ws, subj_name, phase_name
    )

    students_qs = Student.objects.filter(department_id__in=department_ids).exclude(
        enrollment_no=''
    )
    by_enroll: dict[str, list[Student]] = {}
    for s in students_qs:
        key = normalize_enrollment_cell(s.enrollment_no)
        if not key:
            continue
        by_enroll.setdefault(key, []).append(s)

    created = updated = skipped = skipped_no_target = 0
    branch_updates: dict[int, Student] = {}

    phase_cache: dict[int, ExamPhase | None] = {}
    subj_cache: dict[int, Subject | None] = {}

    def phase_for_dept(did: int) -> ExamPhase | None:
        if did not in phase_cache:
            phase_cache[did] = ExamPhase.objects.filter(
                department_id=did, name=phase_name
            ).first()
        return phase_cache[did]

    def subject_for_dept(did: int) -> Subject | None:
        if did not in subj_cache:
            subj_cache[did] = Subject.objects.filter(
                department_id=did, name=subj_name
            ).first()
        return subj_cache[did]

    phase_subject_ok: dict[tuple[int, int], bool] = {}

    def phase_subject_linked(ep_id: int, su_id: int) -> bool:
        k = (ep_id, su_id)
        if k not in phase_subject_ok:
            phase_subject_ok[k] = ExamPhaseSubject.objects.filter(
                exam_phase_id=ep_id, subject_id=su_id
            ).exists()
        return phase_subject_ok[k]

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        if row_idx <= data_start_row or not row:
            continue
        enroll_raw = row[enroll_col] if enroll_col < len(row) else None
        ekey = normalize_enrollment_cell(enroll_raw)
        if not ekey:
            continue
        candidates = by_enroll.get(ekey)
        if not candidates:
            skipped += 1
            continue
        if len(candidates) > 1:
            skipped += 1
            continue
        student = candidates[0]

        phase = phase_for_dept(student.department_id)
        subject = subject_for_dept(student.department_id)
        if not phase or not subject:
            skipped_no_target += 1
            continue
        if not phase_subject_linked(phase.id, subject.id):
            skipped_no_target += 1
            continue

        marks_val = row[marks_col] if marks_col < len(row) else None
        marks_decimal = normalize_student_mark(marks_val)

        if branch_col is not None and branch_col < len(row):
            braw = row[branch_col]
            bval = str(braw).strip() if braw is not None else ''
            if bval and not (student.branch or '').strip():
                student.branch = bval[:BRANCH_MAX_LEN]
                branch_updates[student.pk] = student

        obj, created_flag = StudentMark.objects.update_or_create(
            student=student,
            exam_phase=phase,
            subject=subject,
            defaults={'marks_obtained': marks_decimal},
        )
        if created_flag:
            created += 1
        else:
            updated += 1

    if branch_updates:
        Student.objects.bulk_update(branch_updates.values(), ['branch'], batch_size=300)

    return {
        'created': created,
        'updated': updated,
        'skipped': skipped,
        'skipped_no_target': skipped_no_target,
        'branch_filled': len(branch_updates),
    }
