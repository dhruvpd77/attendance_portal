"""
DR facilities: phase/department/faculty supervision views and scoped Excel exports.
Accessible to exam section (+ super admin), department exam parents, and sub-units.
"""
from __future__ import annotations

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db.models import Count, Q, QuerySet
from django.shortcuts import redirect, render
from django.views.decorators.http import require_http_methods

from core.exam_coordination_views import (
    _child_profile,
    _dept_child_only,
    _dept_parent_only,
    _exam_section_only,
    _exam_section_portal_access,
    _is_hub_coordinator,
    _parent_profile,
)
from core.semester_scope import exam_section_working_semester_ids
from core.exam_subunit_scope import phases_for_subunit_prof, subunit_supervision_duty_filter_q
from core.models import Department, Faculty, SupervisionDuty, SupervisionExamPhase
from core.supervision_dr_export import build_supervision_dr_excel

LIST_LIMIT = 2500


def _dr_facilities_access(request) -> bool:
    return (
        _exam_section_portal_access(request)
        or _dept_parent_only(request)
        or _dept_child_only(request)
    )


def _scope(request) -> dict | None:
    if _exam_section_portal_access(request):
        if _exam_section_only(request):
            sids = exam_section_working_semester_ids(request)
            dept_ids = list(
                Department.objects.filter(institute_semester_id__in=sids).values_list('pk', flat=True)
            )
            return {
                'label': 'Exam section (selected semesters)',
                'department_ids_allowed': dept_ids,
                'faculty_department_filter': None,
                'phase_qs': SupervisionExamPhase.objects.filter(
                    institute_semester_id__in=sids,
                ),
                'base_duty_q': Q(phase__institute_semester_id__in=sids)
                | Q(phase__department__institute_semester_id__in=sids),
            }
        return {
            'label': 'Institute-wide',
            'department_ids_allowed': None,
            'faculty_department_filter': None,
            'phase_qs': SupervisionExamPhase.objects.all(),
            'base_duty_q': Q(),
        }
    if _dept_parent_only(request):
        prof = _parent_profile(request)
        if not prof:
            return None
        if _is_hub_coordinator(prof):
            sid = prof.institute_semester_id
            return {
                'label': f'Hub ({request.user.username}) — {prof.institute_semester.label}',
                'department_ids_allowed': None,
                'faculty_department_filter': None,
                'phase_qs': SupervisionExamPhase.objects.filter(
                    hub_coordinator=request.user,
                    institute_semester_id=sid,
                ),
                'base_duty_q': Q(
                    phase__hub_coordinator=request.user,
                    phase__institute_semester_id=sid,
                ),
            }
        if not prof.department_id:
            return None
        dept = prof.department
        return {
            'label': dept.name,
            'department_ids_allowed': [prof.department_id],
            'faculty_department_filter': prof.department_id,
            'phase_qs': SupervisionExamPhase.objects.filter(department=dept),
            'base_duty_q': Q(phase__department=dept),
        }
    if _dept_child_only(request):
        prof = _child_profile(request)
        if not prof:
            return None
        code = (prof.subunit_code or '').strip().upper()
        dept = prof.department
        return {
            'label': f'{dept.name} ({code})',
            'department_ids_allowed': [prof.department_id],
            'faculty_department_filter': prof.department_id,
            'phase_qs': phases_for_subunit_prof(prof),
            'base_duty_q': subunit_supervision_duty_filter_q(prof),
        }
    return None


def _department_choices(scope: dict):
    allowed = scope['department_ids_allowed']
    if allowed is None:
        return Department.objects.order_by('name')
    return Department.objects.filter(pk__in=allowed).select_related('institute_semester').order_by(
        'institute_semester_id', 'name'
    )


def _faculty_choices(scope: dict):
    qs = Faculty.objects.select_related('department').order_by('department__name', 'full_name')
    fd = scope['faculty_department_filter']
    if fd is not None:
        qs = qs.filter(department_id=fd)
    else:
        allowed = scope['department_ids_allowed']
        if allowed is not None:
            qs = qs.filter(department_id__in=allowed)
    return qs


def _parse_int(val) -> int | None:
    try:
        return int(val)
    except (TypeError, ValueError):
        return None


def _allowed_department(scope, dept_id: int | None) -> bool:
    if dept_id is None:
        return True
    allowed = scope['department_ids_allowed']
    if allowed is None:
        return True
    return dept_id in allowed


def _summary_duties_qs(
    scope: dict,
    phase_id: int | None = None,
    department_id: int | None = None,
    faculty_id: int | None = None,
) -> QuerySet:
    qs = SupervisionDuty.objects.filter(scope['base_duty_q'])
    if phase_id:
        qs = qs.filter(phase_id=phase_id)
    if department_id and _allowed_department(scope, department_id):
        qs = qs.filter(faculty__department_id=department_id)
    if faculty_id:
        fac = Faculty.objects.filter(pk=faculty_id).first()
        if fac:
            fd = scope['faculty_department_filter']
            allowed = scope['department_ids_allowed']
            ok = fac.department_id == fd if fd is not None else (
                allowed is None or fac.department_id in allowed
            )
            if ok:
                qs = qs.filter(faculty_id=faculty_id)
    return qs


def _filtered_duties(
    request,
    scope: dict,
    *,
    phase_id: int | None = None,
    department_id: int | None = None,
    faculty_id: int | None = None,
) -> QuerySet:
    return _summary_duties_qs(scope, phase_id, department_id, faculty_id).select_related(
        'phase', 'faculty', 'faculty__department', 'original_faculty'
    ).order_by('supervision_date', 'phase__name', 'faculty__full_name', 'time_slot')


@login_required
def dr_facilities_dashboard(request):
    if not _dr_facilities_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    scope = _scope(request)
    if not scope:
        messages.error(request, 'No supervision scope for your account.')
        return redirect('accounts:role_redirect')

    phase_id = _parse_int(request.GET.get('phase_id'))
    department_id = _parse_int(request.GET.get('department_id'))
    faculty_id = _parse_int(request.GET.get('faculty_id'))

    if department_id and not _allowed_department(scope, department_id):
        department_id = None

    phases = list(scope['phase_qs'].order_by('name'))
    phase_ids = {p.pk for p in phases}
    if phase_id and phase_id not in phase_ids:
        phase_id = None

    dept_choices = list(_department_choices(scope))
    fac_choices = list(_faculty_choices(scope))

    duties_qs = _filtered_duties(
        request,
        scope,
        phase_id=phase_id,
        department_id=department_id,
        faculty_id=faculty_id,
    )
    total_count = duties_qs.count()
    duties = list(duties_qs[:LIST_LIMIT])

    sum_qs = _summary_duties_qs(scope, phase_id, department_id, faculty_id)
    by_phase = list(
        sum_qs.values('phase__name', 'phase_id').annotate(n=Count('id')).order_by('phase__name')
    )
    by_dept = list(
        sum_qs.values('faculty__department__name', 'faculty__department_id')
        .annotate(n=Count('id'))
        .order_by('faculty__department__name')
    )

    context = {
        'scope_label': scope['label'],
        'show_institute_filters': scope['department_ids_allowed'] is None
        or _exam_section_only(request),
        'phases': phases,
        'departments': dept_choices,
        'faculties': fac_choices,
        'phase_id': phase_id,
        'department_id': department_id,
        'faculty_id': faculty_id,
        'duties': duties,
        'total_count': total_count,
        'truncated': total_count > LIST_LIMIT,
        'by_phase': by_phase,
        'by_dept': by_dept,
        'is_exam_section_view': _exam_section_portal_access(request),
    }
    return render(request, 'core/dr_facilities/dashboard.html', context)


@login_required
@require_http_methods(['GET'])
def dr_facilities_export_department(request):
    if not _dr_facilities_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    scope = _scope(request)
    if not scope:
        return redirect('core:dr_facilities_dashboard')
    dept_id = _parse_int(request.GET.get('department_id'))
    phase_id = _parse_int(request.GET.get('phase_id'))
    if not dept_id or not _allowed_department(scope, dept_id):
        messages.error(request, 'Choose a valid department.')
        return redirect('core:dr_facilities_dashboard')
    dept = Department.objects.filter(pk=dept_id).first()
    if not dept:
        return redirect('core:dr_facilities_dashboard')
    qs = _filtered_duties(request, scope, phase_id=phase_id, department_id=dept_id)
    duties = list(qs)
    title = f'{dept.name} - supervision detail'
    prefix = f'dept_{dept_id}_{phase_id or "all"}'
    return build_supervision_dr_excel(duties, title_line=title, sheet_prefix=prefix)


@login_required
@require_http_methods(['GET'])
def dr_facilities_export_faculty(request):
    if not _dr_facilities_access(request):
        messages.error(request, 'Access denied.')
        return redirect('accounts:role_redirect')
    scope = _scope(request)
    if not scope:
        return redirect('core:dr_facilities_dashboard')
    faculty_id = _parse_int(request.GET.get('faculty_id'))
    phase_id = _parse_int(request.GET.get('phase_id'))
    if not faculty_id:
        messages.error(request, 'Choose a faculty.')
        return redirect('core:dr_facilities_dashboard')
    fac = Faculty.objects.filter(pk=faculty_id).select_related('department').first()
    if not fac:
        return redirect('core:dr_facilities_dashboard')
    if scope['faculty_department_filter'] is not None and fac.department_id != scope['faculty_department_filter']:
        messages.error(request, 'That faculty is outside your access.')
        return redirect('core:dr_facilities_dashboard')
    qs = _filtered_duties(request, scope, phase_id=phase_id, faculty_id=faculty_id)
    duties = list(qs)
    title = f'{fac.full_name} ({fac.short_name}) - supervision'
    prefix = f'fac_{faculty_id}_{phase_id or "all"}'
    return build_supervision_dr_excel(duties, title_line=title, sheet_prefix=prefix)
