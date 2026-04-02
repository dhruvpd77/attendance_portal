"""
Faculty portal: working department in session + multi-department memberships.

Duplicate timetable rows (same person, different ``Department`` / ``Faculty`` PK) are
resolved for **every** faculty user via ``faculty_ids_equivalent_for_portal`` and
``faculty_accessible_departments_qs``: same academic semester + same ``short_name``,
plus either one login, a single "primary" username equal to ``short_name``, or several
``short_name``+digits users (e.g. ``dvp121`` and ``dvp127``) with one full name.
No per-person hardcoding.
"""
from __future__ import annotations

from django.db.models import Q

SESSION_KEY_FACULTY_DEPARTMENT = 'faculty_department_id'


def _username_looks_like_faculty_duplicate_login(username: str, sn_l: str) -> bool:
    """True if username is ``short_name`` or ``short_name`` + digits only (e.g. dvp121)."""
    if not username or not sn_l:
        return False
    u = username.strip().lower()
    if u == sn_l:
        return True
    if not u.startswith(sn_l):
        return False
    rest = u[len(sn_l) :]
    return rest.isdigit() and rest != ''


def portal_faculty_root(faculty):
    """The faculty row that owns the login (never points to portal_canonical)."""
    if not faculty:
        return None
    if faculty.portal_canonical_id:
        return faculty.portal_canonical
    return faculty


def _auto_merged_faculty_pks(faculty) -> set[int] | None:
    """
    When several ``Faculty`` rows share the same academic semester and short name, treat
    them as one person if:

    * exactly one row has a User (duplicate rows without login), **or**
    * more than one row has a User but exactly one is the **primary** account, identified
      by ``User.username`` equal to ``short_name`` (case-insensitive). This covers
      duplicate credentials (e.g. ``mhp`` on SY_4 and ``mhp217`` on SY_3).
    * more than one row has a User and every username looks like ``short_name`` or
      ``short_name``+digits, and all rows share the same ``full_name`` — e.g. ``dvp121``
      on SY_1 and ``dvp127`` on SY_4 for the same superadmin split.

    Returns None if this rule does not apply.
    """
    from core.models import Faculty

    if not faculty or not getattr(faculty, 'department_id', None):
        return None
    sem_id = getattr(faculty.department, 'institute_semester_id', None)
    if not sem_id:
        return None
    sn = (faculty.short_name or '').strip()
    if not sn:
        return None
    sn_l = sn.lower()

    cohort_pks = list(
        Faculty.objects.filter(
            department__institute_semester_id=sem_id,
            short_name__iexact=sn,
        ).values_list('pk', flat=True)
    )
    if len(cohort_pks) < 2:
        return None

    cohort = list(
        Faculty.objects.filter(pk__in=cohort_pks)
        .select_related('user')
        .order_by('pk')
    )
    with_user_rows = [f for f in cohort if f.user_id]
    n_u = len(with_user_rows)
    if n_u == 1:
        return set(cohort_pks)
    if n_u > 1:
        primary = [f for f in with_user_rows if f.user and f.user.username.lower() == sn_l]
        if len(primary) == 1:
            return set(cohort_pks)
        name_keys = {(f.full_name or '').strip().lower() for f in cohort}
        if len(name_keys) == 1 and all(
            f.user and _username_looks_like_faculty_duplicate_login(f.user.username, sn_l)
            for f in with_user_rows
        ):
            return set(cohort_pks)
    return None


def faculty_ids_equivalent_for_portal(portal_faculty) -> set[int]:
    """
    PKs for the same real person: explicit ``portal_canonical`` links, plus automatic
    merge when same institute semester + short name and exactly one login exists.
    """
    from core.models import Faculty

    if not portal_faculty:
        return set()
    root = portal_faculty_root(portal_faculty)
    ids = {root.pk}
    ids.update(
        Faculty.objects.filter(portal_canonical_id=root.pk).values_list('pk', flat=True)
    )
    auto = _auto_merged_faculty_pks(portal_faculty) or _auto_merged_faculty_pks(root)
    if auto:
        ids |= auto
    return ids


def _faculty_portal_department_ids(faculty) -> set[int]:
    from core.models import Faculty, FacultyDepartmentMembership

    if not faculty:
        return set()
    root = portal_faculty_root(faculty)
    ids = set(
        FacultyDepartmentMembership.objects.filter(faculty=root).values_list(
            'department_id', flat=True
        )
    )
    if root.department_id:
        ids.add(root.department_id)
    ids.update(
        Faculty.objects.filter(portal_canonical_id=root.pk).values_list(
            'department_id', flat=True
        )
    )
    equiv = faculty_ids_equivalent_for_portal(faculty)
    ids.update(Faculty.objects.filter(pk__in=equiv).values_list('department_id', flat=True))
    return ids


def faculty_portal_member_departments_qs(faculty):
    """Departments with master portal switch on — any academic semester (attendance may still be unavailable if semester portal is off)."""
    from core.models import Department

    ids = _faculty_portal_department_ids(faculty)
    if not ids:
        return Department.objects.none()
    return (
        Department.objects.filter(pk__in=ids, faculty_portal_enabled=True)
        .select_related('institute_semester')
        .order_by('-institute_semester__sort_order', 'institute_semester_id', 'name')
    )


def faculty_accessible_departments_qs(faculty):
    """Departments for attendance / marks / DR — only when academic semester Faculty portal is open."""
    return faculty_portal_member_departments_qs(faculty).filter(
        institute_semester__faculty_portal_active=True
    )


def faculty_exam_eligible_departments_qs(faculty):
    """Departments where exam menus may appear (ignores semester Faculty portal flag)."""
    from django.db.models import Q

    return faculty_portal_member_departments_qs(faculty).filter(
        Q(faculty_show_exam_duties=True) | Q(faculty_show_exam_credits_analytics=True)
    )


def get_faculty_working_department(request):
    """Department context for attendance & closed-semester features: open academic semester only."""
    if not getattr(request, 'user', None) or not request.user.is_authenticated:
        return None
    faculty = getattr(request.user, 'faculty_profile', None)
    if not faculty:
        return None
    qs = faculty_accessible_departments_qs(faculty)
    if not qs.exists():
        return None
    raw = request.session.get(SESSION_KEY_FACULTY_DEPARTMENT)
    if raw is not None:
        try:
            did = int(raw)
        except (TypeError, ValueError):
            did = None
        if did:
            d = qs.filter(pk=did).first()
            if d:
                return d
    root = portal_faculty_root(faculty)
    if root.department_id:
        d = qs.filter(pk=root.department_id).first()
        if d:
            return d
    return qs.first()


def get_faculty_exam_context_department(request):
    """Department used for exam duties / credits toggles and credit rules (includes closed-semester divisions)."""
    if not getattr(request, 'user', None) or not request.user.is_authenticated:
        return None
    faculty = getattr(request.user, 'faculty_profile', None)
    if not faculty:
        return None
    qs = faculty_exam_eligible_departments_qs(faculty)
    if not qs.exists():
        return None
    raw = request.session.get(SESSION_KEY_FACULTY_DEPARTMENT)
    if raw is not None:
        try:
            did = int(raw)
        except (TypeError, ValueError):
            did = None
        if did:
            d = qs.filter(pk=did).first()
            if d:
                return d
    root = portal_faculty_root(faculty)
    if root.department_id:
        d = qs.filter(pk=root.department_id).first()
        if d:
            return d
    return qs.first()


def faculty_has_department_access(faculty, dept) -> bool:
    if not faculty or not dept:
        return False
    root = portal_faculty_root(faculty)
    if root.department_id == dept.id:
        return True
    from core.models import Faculty, FacultyDepartmentMembership

    if FacultyDepartmentMembership.objects.filter(faculty=root, department=dept).exists():
        return True
    if Faculty.objects.filter(portal_canonical_id=root.pk, department_id=dept.id).exists():
        return True
    equiv = faculty_ids_equivalent_for_portal(faculty)
    return (
        Faculty.objects.filter(pk__in=equiv, department_id=dept.id).exists()
    )


def set_faculty_working_department(request, department_id: int) -> bool:
    """Persist selected department if allowed (master portal on — may be closed-semester for exam context)."""
    faculty = getattr(request.user, 'faculty_profile', None)
    if not faculty:
        return False
    d = faculty_portal_member_departments_qs(faculty).filter(pk=department_id).first()
    if not d:
        return False
    request.session[SESSION_KEY_FACULTY_DEPARTMENT] = d.pk
    return True


def faculty_queryset_for_department_listing(dept):
    """Faculties who belong to this department (primary or extra membership)."""
    from core.models import Faculty

    if not dept:
        return Faculty.objects.none()
    return (
        Faculty.objects.filter(Q(department=dept) | Q(department_memberships__department=dept))
        .select_related('department')
        .distinct()
        .order_by('full_name')
    )
