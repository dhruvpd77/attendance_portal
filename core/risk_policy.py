"""Dynamic at-risk thresholds: attendance % (all phases) and marks per exam phase."""
from __future__ import annotations

from decimal import Decimal

from django.db.models import Q

DEFAULT_ATTENDANCE_MIN_PCT = Decimal('75')
DEFAULT_MARK_FAIL_BELOW = Decimal('9')


def attendance_risk_min_percent(dept) -> Decimal:
    """Cumulative attendance below this percentage counts as at-risk (applies across T1–T4)."""
    if not dept:
        return DEFAULT_ATTENDANCE_MIN_PCT
    if getattr(dept, 'risk_attendance_min_percent', None) is not None:
        return Decimal(str(dept.risk_attendance_min_percent))
    sem = getattr(dept, 'institute_semester', None)
    if sem is not None and getattr(sem, 'risk_attendance_min_percent', None) is not None:
        return Decimal(str(sem.risk_attendance_min_percent))
    return DEFAULT_ATTENDANCE_MIN_PCT


def mark_fail_below_threshold(dept, phase_name: str) -> Decimal:
    """Marks strictly below this value are treated as failed / at-risk for that phase."""
    from core.models import DepartmentExamPhaseRiskThreshold, InstituteExamPhaseRiskThreshold

    pn = (phase_name or '').strip()
    if not pn or not dept:
        return DEFAULT_MARK_FAIL_BELOW
    dr = (
        DepartmentExamPhaseRiskThreshold.objects.filter(department=dept)
        .filter(Q(phase_name__iexact=pn))
        .first()
    )
    if dr:
        return Decimal(str(dr.fail_below_marks))
    sem = getattr(dept, 'institute_semester', None)
    if sem:
        ir = (
            InstituteExamPhaseRiskThreshold.objects.filter(institute_semester=sem)
            .filter(Q(phase_name__iexact=pn))
            .first()
        )
        if ir:
            return Decimal(str(ir.fail_below_marks))
    return DEFAULT_MARK_FAIL_BELOW


def mark_fail_below_getter_for_departments(depts):
    """
    Bulk-load thresholds for many departments; return (dept, phase_name) -> float.

    Use in large Excel exports instead of mark_fail_below_threshold, which runs
    up to two queries per call.
    """
    default = float(DEFAULT_MARK_FAIL_BELOW)
    if not depts:
        return lambda dept, phase_name: default

    from core.models import DepartmentExamPhaseRiskThreshold, InstituteExamPhaseRiskThreshold

    dept_ids = [d.id for d in depts if getattr(d, 'id', None)]
    dept_level = {}
    if dept_ids:
        for t in DepartmentExamPhaseRiskThreshold.objects.filter(department_id__in=dept_ids):
            pn = (t.phase_name or '').strip().lower()
            if pn:
                dept_level[(t.department_id, pn)] = float(Decimal(str(t.fail_below_marks)))

    sem_ids = set()
    for d in depts:
        sid = getattr(d, 'institute_semester_id', None)
        if sid:
            sem_ids.add(sid)

    inst_level = {}
    if sem_ids:
        for t in InstituteExamPhaseRiskThreshold.objects.filter(institute_semester_id__in=sem_ids):
            pn = (t.phase_name or '').strip().lower()
            if pn:
                inst_level[(t.institute_semester_id, pn)] = float(Decimal(str(t.fail_below_marks)))

    def get_lb(dept, phase_name: str) -> float:
        if not dept:
            return default
        pn = (phase_name or '').strip().lower()
        if not pn:
            return default
        did = getattr(dept, 'id', None)
        if did is not None:
            k = (did, pn)
            if k in dept_level:
                return dept_level[k]
        sid = getattr(dept, 'institute_semester_id', None)
        if sid is not None:
            ik = (sid, pn)
            if ik in inst_level:
                return inst_level[ik]
        return default

    return get_lb


def phase_names_for_institute_semester(sem) -> list[str]:
    from core.models import ExamPhase
    from core.exam_phase_order import exam_phase_name_sort_key

    if not sem:
        return []
    names = ExamPhase.objects.filter(department__institute_semester=sem).values_list('name', flat=True).distinct()
    return sorted({str(n).strip() for n in names if n}, key=exam_phase_name_sort_key)
