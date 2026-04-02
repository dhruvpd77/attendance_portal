"""Paper setting credits from DepartmentExamCreditRule and Daily Report column helpers."""
from __future__ import annotations

from collections import defaultdict
from datetime import date
from decimal import Decimal

from django.conf import settings

from core.models import DepartmentExamCreditRule, PaperSettingCompletionRequest


def paper_setting_phase_bucket(phase_name: str) -> str:
    n = (phase_name or '').upper()
    if 'FAST' in n or ' FT' in f' {n}' or n.startswith('FT'):
        return DepartmentExamCreditRule.BUCKET_FAST_TRACK
    if 'REM' in n:
        return DepartmentExamCreditRule.BUCKET_REMEDIAL
    if 'SEE' in n or 'T4' in n:
        return DepartmentExamCreditRule.BUCKET_SEE
    return DepartmentExamCreditRule.BUCKET_T1_T3


def paper_setting_dr_column_for_bucket(bucket: str) -> int:
    mapping = getattr(settings, 'EXAM_DR_PAPER_SETTING_COLUMNS', None) or {
        DepartmentExamCreditRule.BUCKET_T1_T3: 17,
        DepartmentExamCreditRule.BUCKET_SEE: 18,
        DepartmentExamCreditRule.BUCKET_REMEDIAL: 19,
        DepartmentExamCreditRule.BUCKET_FAST_TRACK: 20,
    }
    return int(mapping.get(bucket, 17))


def resolve_exam_credit(
    department_id: int | None,
    task: str,
    bucket: str,
    subject_name: str,
) -> Decimal:
    """Department + subject override → department default → institute subject → institute default."""
    sn = (subject_name or '').strip()

    def _one(qs):
        row = qs.only('credit').first()
        if row is None:
            return None
        return row.credit if row.credit is not None else Decimal('0')

    if department_id and sn:
        v = _one(
            DepartmentExamCreditRule.objects.filter(
                department_id=department_id,
                task=task,
                phase_bucket=bucket,
                subject_name__iexact=sn,
            )
        )
        if v is not None:
            return v
    if department_id:
        v = _one(
            DepartmentExamCreditRule.objects.filter(
                department_id=department_id,
                task=task,
                phase_bucket=bucket,
                subject_name='',
            )
        )
        if v is not None:
            return v
    if sn:
        v = _one(
            DepartmentExamCreditRule.objects.filter(
                department__isnull=True,
                task=task,
                phase_bucket=bucket,
                subject_name__iexact=sn,
            )
        )
        if v is not None:
            return v
    v = _one(
        DepartmentExamCreditRule.objects.filter(
            department__isnull=True,
            task=task,
            phase_bucket=bucket,
            subject_name='',
        )
    )
    if v is not None:
        return v
    return Decimal('0')


def resolve_exam_remuneration(
    department_id: int | None,
    task: str,
    bucket: str,
    subject_name: str,
) -> Decimal:
    """Same resolution order as credits; uses DepartmentExamCreditRule.remuneration (₹)."""
    sn = (subject_name or '').strip()

    def _one(qs):
        row = qs.only('remuneration').first()
        if row is None:
            return None
        return row.remuneration if row.remuneration is not None else Decimal('0')

    if department_id and sn:
        v = _one(
            DepartmentExamCreditRule.objects.filter(
                department_id=department_id,
                task=task,
                phase_bucket=bucket,
                subject_name__iexact=sn,
            )
        )
        if v is not None:
            return v
    if department_id:
        v = _one(
            DepartmentExamCreditRule.objects.filter(
                department_id=department_id,
                task=task,
                phase_bucket=bucket,
                subject_name='',
            )
        )
        if v is not None:
            return v
    if sn:
        v = _one(
            DepartmentExamCreditRule.objects.filter(
                department__isnull=True,
                task=task,
                phase_bucket=bucket,
                subject_name__iexact=sn,
            )
        )
        if v is not None:
            return v
    v = _one(
        DepartmentExamCreditRule.objects.filter(
            department__isnull=True,
            task=task,
            phase_bucket=bucket,
            subject_name='',
        )
    )
    if v is not None:
        return v
    return Decimal('0')


def credit_for_paper_setting_request(req: PaperSettingCompletionRequest) -> Decimal:
    duty = req.duty
    fac = req.faculty
    bucket = paper_setting_phase_bucket(duty.phase.name if duty.phase else '')
    return resolve_exam_credit(
        fac.department_id,
        DepartmentExamCreditRule.TASK_PAPER_SETTING,
        bucket,
        duty.subject_name or '',
    )


def remuneration_for_paper_setting_request(req: PaperSettingCompletionRequest) -> Decimal:
    duty = req.duty
    fac = req.faculty
    bucket = paper_setting_phase_bucket(duty.phase.name if duty.phase else '')
    return resolve_exam_remuneration(
        fac.department_id,
        DepartmentExamCreditRule.TASK_PAPER_SETTING,
        bucket,
        duty.subject_name or '',
    )


def dr_line_for_paper_setting_completion(req: PaperSettingCompletionRequest) -> str:
    duty = req.duty
    ph = duty.phase.name if duty.phase else ''
    sub = (duty.subject_name or '').strip() or '-'
    cr = credit_for_paper_setting_request(req)
    return f'{sub} — paper setting — {ph} — {cr} cr (approved)'


def paper_setting_approved_by_faculty_date(
    report_date: date,
    *,
    faculty_id_filter: set[int] | None = None,
    hub_coordinator_id: int | None = None,
    hub_institute_semester_id: int | None = None,
    duty_phase_semester_ids: list[int] | None = None,
) -> dict[int, dict]:
    qs = PaperSettingCompletionRequest.objects.filter(
        status=PaperSettingCompletionRequest.APPROVED,
        decided_at__date=report_date,
    )
    if hub_coordinator_id is not None:
        qs = qs.filter(duty__phase__hub_coordinator_id=hub_coordinator_id)
        if hub_institute_semester_id is not None:
            qs = qs.filter(duty__phase__institute_semester_id=hub_institute_semester_id)
    if duty_phase_semester_ids:
        from core.semester_scope import q_completion_duty_phase_in_semesters

        qs = qs.filter(q_completion_duty_phase_in_semesters(duty_phase_semester_ids))
    qs = qs.select_related('duty', 'duty__phase', 'faculty')
    out: dict[int, dict] = {}
    for req in qs:
        fid = req.faculty_id
        if faculty_id_filter is not None and fid not in faculty_id_filter:
            continue
        bucket = paper_setting_phase_bucket(req.duty.phase.name if req.duty.phase else '')
        col = paper_setting_dr_column_for_bucket(bucket)
        key = str(col)
        cred = credit_for_paper_setting_request(req)
        ent = out.setdefault(fid, {'lines': [], 'activity_lines': []})
        ent[key] = (ent.get(key) or Decimal('0')) + cred
        ent['lines'].append(dr_line_for_paper_setting_completion(req))
        subj = (req.duty.subject_name or '').strip() or '—'
        ent['activity_lines'].append(f'Paper setting credit — {subj}')
    return out


def aggregate_paper_setting_for_compile(
    dates: list[date],
    faculty_id_filter: set[int] | None,
    *,
    hub_coordinator_id: int | None = None,
    hub_institute_semester_id: int | None = None,
    duty_phase_semester_ids: list[int] | None = None,
) -> dict[int, dict[str, Decimal]]:
    ds = set(dates)
    qs = PaperSettingCompletionRequest.objects.filter(
        status=PaperSettingCompletionRequest.APPROVED,
        decided_at__date__in=ds,
    )
    if hub_coordinator_id is not None:
        qs = qs.filter(duty__phase__hub_coordinator_id=hub_coordinator_id)
        if hub_institute_semester_id is not None:
            qs = qs.filter(duty__phase__institute_semester_id=hub_institute_semester_id)
    if duty_phase_semester_ids:
        from core.semester_scope import q_completion_duty_phase_in_semesters

        qs = qs.filter(q_completion_duty_phase_in_semesters(duty_phase_semester_ids))
    qs = qs.select_related('duty', 'duty__phase', 'faculty')
    buckets = [
        DepartmentExamCreditRule.BUCKET_T1_T3,
        DepartmentExamCreditRule.BUCKET_SEE,
        DepartmentExamCreditRule.BUCKET_REMEDIAL,
        DepartmentExamCreditRule.BUCKET_FAST_TRACK,
    ]
    col_by_b = {b: str(paper_setting_dr_column_for_bucket(b)) for b in buckets}
    out: dict[int, dict[str, Decimal]] = {}
    for req in qs:
        fid = req.faculty_id
        if faculty_id_filter is not None and fid not in faculty_id_filter:
            continue
        ent = out.setdefault(
            fid,
            {
                col_by_b[DepartmentExamCreditRule.BUCKET_T1_T3]: Decimal('0'),
                col_by_b[DepartmentExamCreditRule.BUCKET_SEE]: Decimal('0'),
                col_by_b[DepartmentExamCreditRule.BUCKET_REMEDIAL]: Decimal('0'),
                col_by_b[DepartmentExamCreditRule.BUCKET_FAST_TRACK]: Decimal('0'),
                'total': Decimal('0'),
            },
        )
        b = paper_setting_phase_bucket(req.duty.phase.name if req.duty.phase else '')
        ck = col_by_b.get(b, col_by_b[DepartmentExamCreditRule.BUCKET_T1_T3])
        c = credit_for_paper_setting_request(req)
        ent[ck] = ent[ck] + c
        ent['total'] = ent['total'] + c
    return out


def department_paper_setting_credit_rows(department_id: int) -> list[dict]:
    reqs = (
        PaperSettingCompletionRequest.objects.filter(
            status=PaperSettingCompletionRequest.APPROVED,
            faculty__department_id=department_id,
        )
        .select_related('faculty', 'duty', 'duty__phase')
        .order_by('-decided_at')
    )
    rows = []
    for r in reqs:
        rows.append(
            {
                'req': r,
                'credit': credit_for_paper_setting_request(r),
            }
        )
    return rows


def supervision_credit_for_phase(department_id: int | None, phase_name: str) -> Decimal:
    b = paper_setting_phase_bucket(phase_name)
    return resolve_exam_credit(
        department_id,
        DepartmentExamCreditRule.TASK_SUPERVISION,
        b,
        '',
    )


def supervision_remuneration_for_phase(department_id: int | None, phase_name: str) -> Decimal:
    b = paper_setting_phase_bucket(phase_name)
    return resolve_exam_remuneration(
        department_id,
        DepartmentExamCreditRule.TASK_SUPERVISION,
        b,
        '',
    )
