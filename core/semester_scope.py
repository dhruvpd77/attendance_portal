"""Active academic semester in session (super admin / exam section) + queryset helpers."""
from __future__ import annotations

from django.db.models import Q
from django.shortcuts import redirect

SESSION_KEY_ACTIVE_INSTITUTE_SEMESTER = 'active_institute_semester_id'


def get_active_institute_semester(request):
    """Semester for multi-department admins; persisted in session. Defaults to latest by sort_order."""
    from core.models import InstituteSemester

    if not request.user.is_authenticated:
        return None
    sid = request.session.get(SESSION_KEY_ACTIVE_INSTITUTE_SEMESTER)
    if sid:
        sem = InstituteSemester.objects.filter(pk=sid).first()
        if sem:
            return sem
    sem = InstituteSemester.objects.order_by('-sort_order', '-pk').first()
    if sem:
        request.session[SESSION_KEY_ACTIVE_INSTITUTE_SEMESTER] = sem.pk
    return sem


def set_active_institute_semester(request, semester) -> None:
    request.session[SESSION_KEY_ACTIVE_INSTITUTE_SEMESTER] = semester.pk


def departments_for_institute_semester(semester):
    from core.models import Department

    if not semester:
        return Department.objects.none()
    return Department.objects.filter(institute_semester=semester).order_by('name')


def _is_super_admin_user(request) -> bool:
    if not getattr(request, 'user', None) or not request.user.is_authenticated:
        return False
    if request.user.is_superuser or request.user.is_staff:
        return True
    try:
        rp = request.user.role_profile
        return rp.role == 'admin' and not rp.department_id
    except Exception:
        return True


def departments_for_admin_request(request):
    """Departments visible to current admin: super admin → active semester only; dept admin → own dept."""
    from core.models import Department

    if not request.user.is_authenticated:
        return Department.objects.none()
    if _is_super_admin_user(request):
        return departments_for_institute_semester(get_active_institute_semester(request))
    try:
        rp = request.user.role_profile
        if rp.role in ('admin', 'hod') and rp.department_id:
            return Department.objects.filter(pk=rp.department_id)
    except Exception:
        pass
    return Department.objects.none()


def exam_coordination_uses_semester_request(request) -> bool:
    try:
        role = request.user.role_profile.role
        return role in ('exam_section', 'dept_exam_parent', 'dept_exam_child')
    except Exception:
        return False


SESSION_KEY_EXAM_SECTION_WORKING_SEMESTER_IDS = 'exam_section_working_semester_ids'


def is_exam_section_operator(request) -> bool:
    """Exam section portal login (not super admin)."""
    if not getattr(request, 'user', None) or not request.user.is_authenticated:
        return False
    try:
        return request.user.role_profile.role == 'exam_section'
    except Exception:
        return False


def exam_section_working_semester_ids(request) -> list[int]:
    """
    Academic semester PKs the exam section user wants included institute-wide.
    Defaults once to every defined semester (highest sort_order first) if unset.
    """
    from core.models import InstituteSemester

    if not is_exam_section_operator(request):
        return []
    raw = request.session.get(SESSION_KEY_EXAM_SECTION_WORKING_SEMESTER_IDS)
    if raw is None:
        ids = list(
            InstituteSemester.objects.order_by('-sort_order', '-pk').values_list('pk', flat=True)
        )
        request.session[SESSION_KEY_EXAM_SECTION_WORKING_SEMESTER_IDS] = ids
        return ids
    if not isinstance(raw, list):
        return []
    out = []
    for x in raw:
        try:
            out.append(int(x))
        except (TypeError, ValueError):
            continue
    return out


def set_exam_section_working_semester_ids(request, ids: list[int]) -> None:
    request.session[SESSION_KEY_EXAM_SECTION_WORKING_SEMESTER_IDS] = list(ids)


def departments_for_exam_section_scope(request):
    """Departments limited to exam section working semesters."""
    from core.models import Department

    ids = exam_section_working_semester_ids(request)
    if not ids:
        return Department.objects.none()
    return (
        Department.objects.filter(institute_semester_id__in=ids)
        .select_related('institute_semester')
        .order_by('institute_semester_id', 'name')
    )


def q_supervision_duty_phase_in_semesters(ids: list[int]) -> Q:
    """Supervision duties whose phase belongs to one of these institute semesters."""
    if not ids:
        return Q(pk__isnull=False)
    return Q(phase__institute_semester_id__in=ids) | Q(
        phase__department__institute_semester_id__in=ids
    )


def q_completion_duty_phase_in_semesters(ids: list[int]) -> Q:
    """Paper-check / paper-setting completion rows scoped by duty phase semester."""
    if not ids:
        return Q(pk__isnull=False)
    return (
        Q(duty__phase__institute_scope=True, duty__phase__institute_semester__isnull=True)
        | Q(duty__phase__institute_semester_id__in=ids)
        | Q(duty__phase__department__institute_semester_id__in=ids)
    )


SESSION_KEY_COORD_PARENT_PROFILE_ID = 'coord_parent_exam_profile_id'


def coordinator_parent_profiles_qs(request):
    """Parent coordinator profiles (hub or single-dept) for the logged-in dept_exam_parent user."""
    from core.models import DepartmentExamProfile

    if not request.user.is_authenticated:
        return DepartmentExamProfile.objects.none()
    try:
        if request.user.role_profile.role != 'dept_exam_parent':
            return DepartmentExamProfile.objects.none()
    except Exception:
        return DepartmentExamProfile.objects.none()
    return (
        DepartmentExamProfile.objects.filter(user=request.user, parent__isnull=True)
        .select_related('department', 'institute_semester')
        .order_by('-institute_semester__sort_order', 'institute_semester_id', 'pk')
    )


def get_active_parent_exam_profile(request):
    """Resolved parent DepartmentExamProfile for dept_exam_parent (multi-semester hub uses session)."""
    profiles = list(coordinator_parent_profiles_qs(request))
    if not profiles:
        return None
    if len(profiles) == 1:
        request.session[SESSION_KEY_COORD_PARENT_PROFILE_ID] = profiles[0].pk
        return profiles[0]
    raw = request.session.get(SESSION_KEY_COORD_PARENT_PROFILE_ID)
    pid = None
    if raw is not None:
        try:
            pid = int(raw)
        except (TypeError, ValueError):
            pid = None
    if pid:
        for p in profiles:
            if p.pk == pid:
                return p
    return None


def coordinator_must_select_exam_context(request) -> bool:
    """True when coordinator has multiple parent profiles but none selected in session."""
    profiles = list(coordinator_parent_profiles_qs(request))
    return len(profiles) > 1 and get_active_parent_exam_profile(request) is None


SESSION_KEY_COORD_CHILD_PROFILE_ID = 'coord_child_exam_profile_id'


def coordinator_child_profiles_qs(request):
    """Child (sub-unit) profiles for dept_exam_child users — one row per academic-semester context."""
    from core.models import DepartmentExamProfile

    if not request.user.is_authenticated:
        return DepartmentExamProfile.objects.none()
    try:
        if request.user.role_profile.role != 'dept_exam_child':
            return DepartmentExamProfile.objects.none()
    except Exception:
        return DepartmentExamProfile.objects.none()
    return (
        DepartmentExamProfile.objects.filter(user=request.user, parent__isnull=False)
        .select_related('department', 'institute_semester', 'parent')
        .order_by('-institute_semester__sort_order', 'institute_semester_id', 'pk')
    )


def get_active_child_exam_profile(request):
    """Resolved child DepartmentExamProfile when same sub-unit login is used for more than one semester."""
    profiles = list(coordinator_child_profiles_qs(request))
    if not profiles:
        return None
    if len(profiles) == 1:
        request.session[SESSION_KEY_COORD_CHILD_PROFILE_ID] = profiles[0].pk
        return profiles[0]
    raw = request.session.get(SESSION_KEY_COORD_CHILD_PROFILE_ID)
    pid = None
    if raw is not None:
        try:
            pid = int(raw)
        except (TypeError, ValueError):
            pid = None
    if pid:
        for p in profiles:
            if p.pk == pid:
                return p
    return None


def child_must_select_exam_context(request) -> bool:
    profiles = list(coordinator_child_profiles_qs(request))
    return len(profiles) > 1 and get_active_child_exam_profile(request) is None


def institute_semester_for_exam_portal(request):
    """Semester for exam coordinators: active parent profile, session for hub/exam section; else linked department's semester."""
    from core.models import DepartmentExamProfile

    sem = get_active_institute_semester(request)
    if not exam_coordination_uses_semester_request(request):
        return sem
    try:
        role = request.user.role_profile.role
    except Exception:
        return sem
    if role == 'dept_exam_parent':
        prof = get_active_parent_exam_profile(request)
        if prof:
            if prof.department_id and prof.department.institute_semester_id:
                return prof.department.institute_semester
            if prof.institute_semester_id:
                return prof.institute_semester
        return sem
    if role == 'dept_exam_child':
        prof = get_active_child_exam_profile(request)
        if prof:
            if prof.department_id and prof.department.institute_semester_id:
                return prof.department.institute_semester
            if prof.institute_semester_id:
                return prof.institute_semester
        return sem
    prof = (
        DepartmentExamProfile.objects.filter(user=request.user)
        .select_related('department', 'institute_semester')
        .first()
    )
    if prof and prof.department_id and prof.department.institute_semester_id:
        return prof.department.institute_semester
    if prof and prof.institute_semester_id:
        return prof.institute_semester
    return sem


def _exam_coordination_wide_department_queryset(request):
    """All departments across academic semesters — exam section & super admin need FY/SY pickers."""
    from core.models import Department, DepartmentExamProfile

    if not request.user.is_authenticated:
        return None
    qs = (
        Department.objects.select_related('institute_semester')
        .order_by('-institute_semester__sort_order', 'institute_semester_id', 'name')
    )
    try:
        if request.user.role_profile.role == 'exam_section':
            return None
    except Exception:
        pass
    if _is_super_admin_user(request):
        return qs
    if exam_coordination_uses_semester_request(request):
        try:
            prof = get_active_parent_exam_profile(request)
            if (
                prof
                and prof.parent_id is None
                and prof.is_hub_coordinator
                and not prof.department_id
            ):
                return qs
        except Exception:
            pass
    return None


def departments_for_exam_coordination_request(request):
    """Department pickers: exam section → working semesters; super admin → all semesters; others → one academic semester."""
    if is_exam_section_operator(request):
        return departments_for_exam_section_scope(request)
    wide = _exam_coordination_wide_department_queryset(request)
    if wide is not None:
        return wide
    sem = institute_semester_for_exam_portal(request)
    return departments_for_institute_semester(sem).order_by('name')


def super_admin_must_select_semester_response(request, view_name: str = 'core:admin_semester_list'):
    from core.models import InstituteSemester

    if not _is_super_admin_user(request):
        return None
    if not InstituteSemester.objects.exists():
        return None
    if get_active_institute_semester(request):
        return None
    return redirect(view_name)

