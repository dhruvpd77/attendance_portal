def _is_super_admin(request):
    """True if admin with no department (can create depts and departmental admins)."""
    if not request.user.is_authenticated:
        return False
    if request.user.is_superuser or request.user.is_staff:
        return True
    try:
        rp = request.user.role_profile
        return rp.role == 'admin' and not rp.department_id
    except Exception:
        return False


def _make_link(request, label, url_name, icon, active_name):
    """Build a link dict with resolved URL so sidebar hrefs always work."""
    from django.urls import reverse
    try:
        path = reverse(url_name)
        url = request.build_absolute_uri(path)
    except Exception:
        url = '#'
    return {'label': label, 'url': url, 'icon': icon, 'active_name': active_name}


def sidebar_links(request):
    links = []
    if not request.user.is_authenticated:
        return {'sidebar_links': links, 'is_super_admin': False}
    is_super_admin = False
    try:
        from accounts.models import UserRole
        role = request.user.role_profile.role
    except Exception:
        role = 'admin' if (request.user.is_superuser or request.user.is_staff) else None
    if role == 'admin' or request.user.is_superuser or request.user.is_staff:
        is_super_admin = _is_super_admin(request)
        links = [
            _make_link(request, 'Dashboard', 'core:admin_dashboard', 'fa-tachometer-alt', 'admin_dashboard'),
            _make_link(request, 'Analytics', 'core:admin_analytics_dashboard', 'fa-chart-line', 'admin_analytics_dashboard'),
            _make_link(request, 'Departments', 'core:department_list', 'fa-building', 'department_list'),
            _make_link(request, 'Batches', 'core:batch_list', 'fa-layer-group', 'batch_list'),
            _make_link(request, 'Subjects', 'core:subject_list', 'fa-book', 'subject_list'),
            _make_link(request, 'Faculties', 'core:faculty_list', 'fa-chalkboard-teacher', 'faculty_list'),
            _make_link(request, 'Generate credentials', 'core:generate_credentials_choice', 'fa-key', 'generate_credentials_choice'),
            _make_link(request, 'Students', 'core:student_list', 'fa-user-graduate', 'student_list'),
            _make_link(request, 'Schedule', 'core:schedule_list', 'fa-calendar-alt', 'schedule_list'),
            _make_link(request, 'Upload Timetable', 'core:upload_timetable', 'fa-file-upload', 'upload_timetable'),
            _make_link(request, 'Term Phases', 'core:term_phases', 'fa-calendar-week', 'term_phases'),
            _make_link(request, 'Daily Absent', 'core:daily_absent', 'fa-file-excel', 'daily_absent'),
            _make_link(request, 'Attendance Sheet', 'core:attendance_sheet_manager', 'fa-table', 'attendance_sheet_manager'),
            _make_link(request, 'Lecture Adjustment', 'core:lecture_adjustment', 'fa-exchange-alt', 'lecture_adjustment'),
            _make_link(request, 'Compile Attendance', 'core:compile_attendance', 'fa-file-alt', 'compile_attendance'),
        ]
        if is_super_admin:
            links.insert(2, _make_link(request, 'Departmental admins', 'core:departmental_admin_list', 'fa-user-shield', 'departmental_admin_list'))
    elif role == 'faculty':
        links = [
            _make_link(request, 'Dashboard', 'core:faculty_dashboard', 'fa-tachometer-alt', 'faculty_dashboard'),
            _make_link(request, 'Mark Attendance', 'core:faculty_attendance_entry', 'fa-edit', 'faculty_attendance_entry'),
            _make_link(request, 'Mentorship Students', 'core:faculty_mentorship', 'fa-users', 'faculty_mentorship'),
        ]
    elif role == 'student':
        links = [
            _make_link(request, 'Dashboard', 'core:student_dashboard', 'fa-tachometer-alt', 'student_dashboard'),
            _make_link(request, 'Attendance Analytics', 'core:student_attendance_analytics', 'fa-chart-line', 'student_attendance_analytics'),
        ]
    return {'sidebar_links': links, 'is_super_admin': is_super_admin}
