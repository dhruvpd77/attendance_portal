"""Paper checking credit rules, totals for faculty rows, and Daily Report export helpers."""
from __future__ import annotations

from collections import defaultdict
from datetime import date
from decimal import Decimal

from django.db.models import Prefetch

from core.models import (
    DepartmentExamCreditRule,
    Faculty,
    PaperCheckingAdjustedShare,
    PaperCheckingAllocation,
    PaperCheckingCompletionRequest,
    PaperCheckingSubjectCredit,
)
from core.paper_setting_credits import (
    paper_setting_phase_bucket,
    resolve_exam_credit,
    resolve_exam_remuneration,
)


def get_subject_credit(phase_id: int, subject_name: str) -> PaperCheckingSubjectCredit | None:
    sn = (subject_name or '').strip()
    if not sn:
        return None
    return PaperCheckingSubjectCredit.objects.filter(
        phase_id=phase_id, subject_name__iexact=sn
    ).first()


def eval_credit_column_for_phase(phase_name: str) -> int:
    """Workbook column: 23 = T1–T3, 24 = SEE/T4, 25 = REM (matches exam.xlsx)."""
    n = (phase_name or '').upper()
    if 'REM' in n:
        return 25
    if 'SEE' in n or 'T4' in n:
        return 24
    return 23


def paper_count_for_completion(req: PaperCheckingCompletionRequest) -> int:
    sh = PaperCheckingAdjustedShare.objects.filter(
        duty_id=req.duty_id, faculty_id=req.faculty_id
    ).first()
    if sh:
        return sh.paper_count
    return req.duty.total_students


def credit_for_completion_request(req: PaperCheckingCompletionRequest) -> Decimal:
    duty = req.duty
    papers = paper_count_for_completion(req)
    if papers <= 0:
        return Decimal('0')
    rule = get_subject_credit(duty.phase_id, duty.subject_name)
    if rule:
        if rule.is_practical:
            # One submission: same paper count earns both online and offline components.
            r_on = rule.credit_online_per_paper if rule.credit_online_per_paper is not None else Decimal('0')
            r_off = rule.credit_offline_per_paper if rule.credit_offline_per_paper is not None else Decimal('0')
            return Decimal(papers) * (r_on + r_off)
        return Decimal(papers) * (rule.credit_per_paper_theory or Decimal('0'))
    fac = req.faculty
    bucket = paper_setting_phase_bucket(duty.phase.name if duty.phase else '')
    per_paper = resolve_exam_credit(
        fac.department_id if fac else None,
        DepartmentExamCreditRule.TASK_PAPER_CHECKING,
        bucket,
        duty.subject_name or '',
    )
    return Decimal(papers) * per_paper


def remuneration_for_completion_request(req: PaperCheckingCompletionRequest) -> Decimal:
    """₹ using same rule rows as credit: scale by credits when rule credit > 0, else papers × ₹/paper."""
    cred = credit_for_completion_request(req)
    if cred <= 0:
        return Decimal('0')
    duty = req.duty
    fac = req.faculty
    papers = paper_count_for_completion(req)
    bucket = paper_setting_phase_bucket(duty.phase.name if duty.phase else '')
    dept_id = fac.department_id if fac else None
    subj = duty.subject_name or ''
    rc = resolve_exam_credit(
        dept_id,
        DepartmentExamCreditRule.TASK_PAPER_CHECKING,
        bucket,
        subj,
    )
    rr = resolve_exam_remuneration(
        dept_id,
        DepartmentExamCreditRule.TASK_PAPER_CHECKING,
        bucket,
        subj,
    )
    if rc and rc > 0:
        return cred * (rr / rc)
    if papers and rr:
        return Decimal(papers) * rr
    return Decimal('0')


def dr_activity_line_for_completion(req: PaperCheckingCompletionRequest) -> str:
    """Not-defined-LJU activity column (block - subject - phase - papers)."""
    duty = req.duty
    fac = req.faculty
    papers = paper_count_for_completion(req)
    block = '-'
    allocs = list(duty.allocations.all())
    for a in allocs:
        if fac.department_id and a.department_id == fac.department_id:
            block = (a.block_range or '-').strip() or '-'
            break
    if block == '-' and allocs:
        block = (allocs[0].block_range or '-').strip() or '-'
    ph = duty.phase.name if duty.phase else ''
    sub = (duty.subject_name or '').strip() or '-'
    rule = get_subject_credit(duty.phase_id, duty.subject_name)
    if rule and rule.is_practical:
        return (
            f'{block} - {sub} - {ph} - {papers} papers (practical: online+offline credit, same count)'
        )
    if rule:
        return f'{block} - {sub} - {ph} - {papers} papers evaluated'
    cr = credit_for_completion_request(req)
    return f'{block} - {sub} - {ph} - {papers} papers evaluated (fallback rate → {cr} cr)'


def paper_eval_approved_by_faculty_date(
    report_date: date,
    *,
    faculty_id_filter: set[int] | None = None,
    hub_coordinator_id: int | None = None,
    hub_institute_semester_id: int | None = None,
    duty_phase_semester_ids: list[int] | None = None,
) -> dict[int, dict]:
    """
    Per faculty: sums for DR columns 23–25 and text lines for col 42.
    Uses coordinator approval date (decided_at) as the DR reporting day.
    """
    qs = PaperCheckingCompletionRequest.objects.filter(
        status=PaperCheckingCompletionRequest.APPROVED,
        decided_at__date=report_date,
    )
    if hub_coordinator_id is not None:
        qs = qs.filter(duty__phase__hub_coordinator_id=hub_coordinator_id)
        if hub_institute_semester_id is not None:
            qs = qs.filter(duty__phase__institute_semester_id=hub_institute_semester_id)
    if duty_phase_semester_ids:
        from core.semester_scope import q_completion_duty_phase_in_semesters

        qs = qs.filter(q_completion_duty_phase_in_semesters(duty_phase_semester_ids))
    qs = (
        qs
        .select_related('duty', 'duty__phase', 'faculty')
        .prefetch_related(
            Prefetch(
                'duty__allocations',
                queryset=PaperCheckingAllocation.objects.select_related('department'),
            ),
        )
    )
    out: dict[int, dict] = defaultdict(
        lambda: {'23': Decimal('0'), '24': Decimal('0'), '25': Decimal('0'), 'lines': []}
    )
    for req in qs:
        fid = req.faculty_id
        if faculty_id_filter is not None and fid not in faculty_id_filter:
            continue
        col = eval_credit_column_for_phase(req.duty.phase.name if req.duty.phase else '')
        cred = credit_for_completion_request(req)
        key = str(col)
        ent = out[fid]
        ent[key] = ent[key] + cred
        ent['lines'].append(dr_activity_line_for_completion(req))
    return {k: dict(v) for k, v in out.items()}


def department_approved_paper_credit_rows(department_id: int) -> list[dict]:
    """Faculty in department with summed approved paper-check credits (for coordinator dashboards)."""
    reqs = (
        PaperCheckingCompletionRequest.objects.filter(
            status=PaperCheckingCompletionRequest.APPROVED,
            faculty__department_id=department_id,
        )
        .select_related('faculty', 'duty', 'duty__phase')
        .prefetch_related(
            Prefetch(
                'duty__allocations',
                queryset=PaperCheckingAllocation.objects.select_related('department'),
            ),
        )
    )
    by_fac: dict[int, Decimal] = defaultdict(lambda: Decimal('0'))
    for r in reqs:
        by_fac[r.faculty_id] += credit_for_completion_request(r)
    if not by_fac:
        return []
    fmap = Faculty.objects.in_bulk(by_fac.keys())
    rows = [{'faculty': fmap[fid], 'credits': by_fac[fid]} for fid in by_fac if fid in fmap]
    rows.sort(key=lambda x: x['faculty'].full_name)
    return rows


def aggregate_paper_check_credits_for_compile(
    dates: list[date],
    faculty_id_filter: set[int] | None,
    *,
    hub_coordinator_id: int | None = None,
    hub_institute_semester_id: int | None = None,
    duty_phase_semester_ids: list[int] | None = None,
) -> dict[int, dict[str, Decimal]]:
    """Approved paper-check credits per faculty across days, by answerbook-eval column bucket (23/24/25)."""
    ds = set(dates)
    qs = PaperCheckingCompletionRequest.objects.filter(
        status=PaperCheckingCompletionRequest.APPROVED,
        decided_at__date__in=ds,
    )
    if hub_coordinator_id is not None:
        qs = qs.filter(duty__phase__hub_coordinator_id=hub_coordinator_id)
        if hub_institute_semester_id is not None:
            qs = qs.filter(duty__phase__institute_semester_id=hub_institute_semester_id)
    if duty_phase_semester_ids:
        from core.semester_scope import q_completion_duty_phase_in_semesters

        qs = qs.filter(q_completion_duty_phase_in_semesters(duty_phase_semester_ids))
    qs = (
        qs.select_related('duty', 'duty__phase', 'faculty')
        .prefetch_related(
            Prefetch(
                'duty__allocations',
                queryset=PaperCheckingAllocation.objects.select_related('department'),
            ),
        )
    )
    out: dict[int, dict[str, Decimal]] = {}
    for r in qs:
        fid = r.faculty_id
        if faculty_id_filter is not None and fid not in faculty_id_filter:
            continue
        ent = out.setdefault(
            fid,
            {'23': Decimal('0'), '24': Decimal('0'), '25': Decimal('0'), 'total': Decimal('0')},
        )
        col = str(eval_credit_column_for_phase(r.duty.phase.name if r.duty.phase else ''))
        c = credit_for_completion_request(r)
        ent[col] = ent[col] + c
        ent['total'] = ent['total'] + c
    return out
