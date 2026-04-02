"""Shared queries for paper duty phases (no view imports — avoids circular deps)."""
from __future__ import annotations

from django.db.models import Prefetch, Q

from core.semester_scope import (
    departments_for_exam_section_scope,
    exam_section_working_semester_ids,
    is_exam_section_operator,
)

from core.models import (
    DepartmentExamProfile,
    Faculty,
    PaperCheckingAdjustedShare,
    PaperCheckingAllocation,
    PaperCheckingCompletionRequest,
    PaperCheckingDuty,
    PaperCheckingPhase,
    PaperSettingCompletionRequest,
    PaperSettingDuty,
    PaperSettingPhase,
)
from core.paper_checking_credits import credit_for_completion_request
from core.paper_setting_credits import credit_for_paper_setting_request


def checking_phases_institute():
    return PaperCheckingPhase.objects.filter(institute_scope=True).order_by('name')


def checking_phases_exam_section_scoped(request):
    ids = exam_section_working_semester_ids(request)
    if not ids:
        return PaperCheckingPhase.objects.none()
    dept_ids = departments_for_exam_section_scope(request).values_list('pk', flat=True)
    return (
        PaperCheckingPhase.objects.filter(
            Q(institute_scope=True)
            & (Q(institute_semester__isnull=True) | Q(institute_semester_id__in=ids))
            | Q(department_id__in=dept_ids)
            | Q(hub_coordinator__isnull=False, institute_semester_id__in=ids)
        )
        .distinct()
        .order_by('name')
    )


def setting_phases_exam_section_scoped(request):
    ids = exam_section_working_semester_ids(request)
    if not ids:
        return PaperSettingPhase.objects.none()
    dept_ids = departments_for_exam_section_scope(request).values_list('pk', flat=True)
    return (
        PaperSettingPhase.objects.filter(
            Q(institute_scope=True)
            & (Q(institute_semester__isnull=True) | Q(institute_semester_id__in=ids))
            | Q(department_id__in=dept_ids)
            | Q(hub_coordinator__isnull=False, institute_semester_id__in=ids)
        )
        .distinct()
        .order_by('name')
    )


def checking_phases_institute_for_request(request):
    if is_exam_section_operator(request):
        return checking_phases_exam_section_scoped(request)
    return PaperCheckingPhase.objects.filter(institute_scope=True).order_by('name')


def checking_phases_hub_user(user, institute_semester_id: int | None = None):
    qs = PaperCheckingPhase.objects.filter(hub_coordinator=user)
    if institute_semester_id:
        qs = qs.filter(institute_semester_id=institute_semester_id)
    return qs.order_by('name')


def checking_phases_department(dept):
    return PaperCheckingPhase.objects.filter(department=dept).order_by('name')


def setting_phases_institute():
    return PaperSettingPhase.objects.filter(institute_scope=True).order_by('name')


def setting_phases_institute_for_request(request):
    if is_exam_section_operator(request):
        return setting_phases_exam_section_scoped(request)
    return PaperSettingPhase.objects.filter(institute_scope=True).order_by('name')


def setting_phases_hub_user(user, institute_semester_id: int | None = None):
    qs = PaperSettingPhase.objects.filter(hub_coordinator=user)
    if institute_semester_id:
        qs = qs.filter(institute_semester_id=institute_semester_id)
    return qs.order_by('name')


def setting_phases_department(dept):
    return PaperSettingPhase.objects.filter(department=dept).order_by('name')


def paper_checking_duties_for_child_prof(prof):
    """Duties where evaluators are assigned blocks for this sub-unit's attendance department only."""
    dept_id = prof.department_id
    parent = prof.parent
    hub_uid = parent.user_id if parent and not parent.department_id else None
    q_phase = Q(department_id=dept_id) | Q(institute_scope=True)
    if hub_uid:
        q_phase |= Q(hub_coordinator_id=hub_uid)
    phase_ids = PaperCheckingPhase.objects.filter(q_phase).values_list('id', flat=True)
    return (
        PaperCheckingDuty.objects.filter(phase_id__in=phase_ids)
        .filter(allocations__department_id=dept_id)
        .select_related('phase', 'faculty')
        .prefetch_related('allocations__department')
        .distinct()
        .order_by('exam_date', 'subject_name')
    )


def pending_paper_check_completion_requests_for_child(prof):
    base_ids = paper_checking_duties_for_child_prof(prof).values_list('id', flat=True)
    return (
        PaperCheckingCompletionRequest.objects.filter(
            duty_id__in=base_ids,
            status=PaperCheckingCompletionRequest.PENDING,
        )
        .select_related('duty', 'duty__phase', 'faculty', 'faculty__department')
        .prefetch_related(
            Prefetch(
                'duty__adjusted_shares',
                queryset=PaperCheckingAdjustedShare.objects.select_related('faculty'),
            ),
        )
        .order_by('submitted_at')
    )


def history_paper_check_completion_requests_for_child(prof, limit: int = 250):
    """Accepted (approved) and dismissed completion requests for this sub-unit’s scope."""
    base_ids = paper_checking_duties_for_child_prof(prof).values_list('id', flat=True)
    return (
        PaperCheckingCompletionRequest.objects.filter(
            duty_id__in=base_ids,
            status__in=(
                PaperCheckingCompletionRequest.APPROVED,
                PaperCheckingCompletionRequest.REJECTED,
            ),
        )
        .select_related('duty', 'duty__phase', 'faculty', 'decided_by')
        .prefetch_related(
            Prefetch(
                'duty__adjusted_shares',
                queryset=PaperCheckingAdjustedShare.objects.select_related('faculty'),
            ),
        )
        .order_by('-decided_at', '-submitted_at', '-id')[:limit]
    )


def _paper_row_approved_credit(faculty: Faculty, duty: PaperCheckingDuty):
    req = (
        PaperCheckingCompletionRequest.objects.filter(
            duty_id=duty.pk, faculty=faculty, status=PaperCheckingCompletionRequest.APPROVED
        )
        .select_related('duty', 'duty__phase')
        .prefetch_related(
            Prefetch(
                'duty__allocations',
                queryset=PaperCheckingAllocation.objects.select_related('department'),
            ),
        )
        .order_by('-decided_at', '-id')
        .first()
    )
    return credit_for_completion_request(req) if req else None


def build_faculty_paper_checking_rows(faculty: Faculty) -> list[dict]:
    """Rows for faculty portal: original duties (when duty not split) + adjusted shares with paper_count and badge."""
    duty_ids_split = PaperCheckingAdjustedShare.objects.values_list('duty_id', flat=True).distinct()
    rows: list[dict] = []
    for share in (
        PaperCheckingAdjustedShare.objects.filter(faculty=faculty)
        .select_related('duty', 'duty__phase')
        .prefetch_related('duty__allocations__department')
        .order_by('duty__exam_date', 'duty__subject_name')
    ):
        rows.append(
            {
                'duty': share.duty,
                'paper_count': share.paper_count,
                'adjusted': True,
                'status': paper_checking_completion_ui_status(faculty, share.duty_id),
                'approved_credit': _paper_row_approved_credit(faculty, share.duty),
            }
        )
    for d in (
        PaperCheckingDuty.objects.filter(faculty=faculty)
        .exclude(pk__in=duty_ids_split)
        .select_related('phase')
        .prefetch_related('allocations__department')
        .order_by('exam_date', 'subject_name')
    ):
        rows.append(
            {
                'duty': d,
                'paper_count': d.total_students,
                'adjusted': False,
                'status': paper_checking_completion_ui_status(faculty, d.pk),
                'approved_credit': _paper_row_approved_credit(faculty, d),
            }
        )
    rows.sort(key=lambda r: (r['duty'].exam_date, (r['duty'].subject_name or '').upper()))
    return rows


def paper_checking_completion_ui_status(faculty, duty_id: int) -> str:
    r = (
        PaperCheckingCompletionRequest.objects.filter(duty_id=duty_id, faculty=faculty)
        .order_by('-submitted_at', '-id')
        .first()
    )
    if not r:
        return 'open'
    if r.status == PaperCheckingCompletionRequest.APPROVED:
        return 'approved'
    if r.status == PaperCheckingCompletionRequest.PENDING:
        return 'pending'
    return 'open'


def paper_setting_duties_for_child_prof(prof):
    dept_id = prof.department_id
    parent = prof.parent
    hub_uid = parent.user_id if parent and not parent.department_id else None
    q_phase = Q(department_id=dept_id) | Q(institute_scope=True)
    if hub_uid:
        q_phase |= Q(hub_coordinator_id=hub_uid)
    phase_ids = PaperSettingPhase.objects.filter(q_phase).values_list('id', flat=True)
    return (
        PaperSettingDuty.objects.filter(phase_id__in=phase_ids, faculty__department_id=dept_id)
        .select_related('phase', 'faculty')
        .order_by('duty_date', 'subject_name')
    )


def hub_managed_department_ids(prof) -> set[int]:
    """Departments invited or with sub-units under this hub coordinator profile."""
    d_inv = set(
        DepartmentExamProfile.objects.filter(
            invited_by=prof.user,
            department_id__isnull=False,
            institute_semester_id=prof.institute_semester_id,
        ).values_list('department_id', flat=True)
    )
    d_ch = set(
        DepartmentExamProfile.objects.filter(parent=prof, department_id__isnull=False).values_list(
            'department_id', flat=True
        )
    )
    return d_inv | d_ch


def pending_paper_setting_completion_requests_for_child(prof):
    base_ids = paper_setting_duties_for_child_prof(prof).values_list('id', flat=True)
    return (
        PaperSettingCompletionRequest.objects.filter(
            duty_id__in=base_ids,
            status=PaperSettingCompletionRequest.PENDING,
        )
        .select_related('duty', 'duty__phase', 'faculty', 'faculty__department')
        .order_by('submitted_at')
    )


def history_paper_setting_completion_requests_for_child(prof, limit: int = 250):
    base_ids = paper_setting_duties_for_child_prof(prof).values_list('id', flat=True)
    return (
        PaperSettingCompletionRequest.objects.filter(
            duty_id__in=base_ids,
            status__in=(
                PaperSettingCompletionRequest.APPROVED,
                PaperSettingCompletionRequest.REJECTED,
            ),
        )
        .select_related('duty', 'duty__phase', 'faculty', 'decided_by')
        .order_by('-decided_at', '-pk')[:limit]
    )


def pending_paper_setting_completion_requests_for_parent(prof):
    hub = bool(prof.parent_id is None and getattr(prof, 'is_hub_coordinator', False))
    if hub:
        dept_ids = hub_managed_department_ids(prof)
        scope = Q(
            duty__phase__hub_coordinator_id=prof.user_id,
            duty__phase__institute_semester_id=prof.institute_semester_id,
        )
        if dept_ids:
            scope |= Q(faculty__department_id__in=dept_ids)
        return (
            PaperSettingCompletionRequest.objects.filter(
                status=PaperSettingCompletionRequest.PENDING,
            )
            .filter(scope)
            .select_related('duty', 'duty__phase', 'faculty', 'faculty__department')
            .distinct()
            .order_by('submitted_at')
        )
    if not prof.department_id:
        return PaperSettingCompletionRequest.objects.none()
    return (
        PaperSettingCompletionRequest.objects.filter(
            status=PaperSettingCompletionRequest.PENDING,
            faculty__department_id=prof.department_id,
        )
        .select_related('duty', 'duty__phase', 'faculty', 'faculty__department')
        .order_by('submitted_at')
    )


def history_paper_setting_completion_requests_for_parent(prof, limit: int = 200):
    hub = bool(prof.parent_id is None and getattr(prof, 'is_hub_coordinator', False))
    qs = PaperSettingCompletionRequest.objects.filter(
        status__in=(
            PaperSettingCompletionRequest.APPROVED,
            PaperSettingCompletionRequest.REJECTED,
        )
    ).select_related('duty', 'duty__phase', 'faculty', 'decided_by')
    if hub:
        dept_ids = hub_managed_department_ids(prof)
        scope = Q(
            duty__phase__hub_coordinator_id=prof.user_id,
            duty__phase__institute_semester_id=prof.institute_semester_id,
        )
        if dept_ids:
            scope |= Q(faculty__department_id__in=dept_ids)
        qs = qs.filter(scope)
    elif prof.department_id:
        qs = qs.filter(faculty__department_id=prof.department_id)
    else:
        return []
    return list(qs.order_by('-decided_at', '-pk')[:limit])


def paper_setting_completion_ui_status(faculty, duty_id: int) -> str:
    r = (
        PaperSettingCompletionRequest.objects.filter(duty_id=duty_id, faculty=faculty)
        .order_by('-submitted_at', '-id')
        .first()
    )
    if not r:
        return 'open'
    if r.status == PaperSettingCompletionRequest.APPROVED:
        return 'approved'
    if r.status == PaperSettingCompletionRequest.PENDING:
        return 'pending'
    return 'open'


def _paper_setting_row_approved_credit(faculty, duty_id: int):
    r = (
        PaperSettingCompletionRequest.objects.filter(
            duty_id=duty_id,
            faculty=faculty,
            status=PaperSettingCompletionRequest.APPROVED,
        )
        .order_by('-decided_at')
        .first()
    )
    if r:
        return credit_for_paper_setting_request(r)
    return None


def build_faculty_paper_setting_rows(faculty):
    rows: list[dict] = []
    for d in (
        PaperSettingDuty.objects.filter(faculty=faculty)
        .select_related('phase')
        .order_by('duty_date', 'subject_name')
    ):
        rows.append(
            {
                'duty': d,
                'status': paper_setting_completion_ui_status(faculty, d.pk),
                'approved_credit': _paper_setting_row_approved_credit(faculty, d.pk),
            }
        )
    return rows
