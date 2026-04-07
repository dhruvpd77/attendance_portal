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


def phase_names_for_institute_semester(sem) -> list[str]:
    from core.models import ExamPhase
    from core.exam_phase_order import exam_phase_name_sort_key

    if not sem:
        return []
    names = ExamPhase.objects.filter(department__institute_semester=sem).values_list('name', flat=True).distinct()
    return sorted({str(n).strip() for n in names if n}, key=exam_phase_name_sort_key)
