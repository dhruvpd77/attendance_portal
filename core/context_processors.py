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


def _make_link(request, label, url_name, icon, active_name, *, active_name_extra=None):
    """Build a link dict with resolved URL so sidebar hrefs always work."""
    from django.urls import reverse
    try:
        path = reverse(url_name)
        url = request.build_absolute_uri(path)
    except Exception:
        url = '#'
    item = {'label': label, 'url': url, 'icon': icon, 'active_name': active_name}
    if active_name_extra:
        item['active_name_extra'] = list(active_name_extra)
    return item


def sidebar_links(request):
    links = []
    empty = {'sidebar_links': links, 'is_super_admin': False, 'faculty_portal_flags': None}
    if not request.user.is_authenticated:
        return empty
    is_super_admin = False
    try:
        from accounts.models import UserRole
        role = request.user.role_profile.role
    except Exception:
        role = 'admin' if (request.user.is_superuser or request.user.is_staff) else None
    if role in ('admin', 'hod') or request.user.is_superuser or request.user.is_staff:
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
            _make_link(request, 'Lock Attendance Time', 'core:attendance_lock_setting', 'fa-lock', 'attendance_lock_setting'),
            _make_link(request, 'Lecture Cancellation', 'core:lecture_cancellation', 'fa-times-circle', 'lecture_cancellation'),
            _make_link(request, 'Add Extra Lecture', 'core:extra_lecture', 'fa-plus-circle', 'extra_lecture'),
            _make_link(request, 'Manual Attendance', 'core:admin_manual_attendance', 'fa-edit', 'admin_manual_attendance'),
            _make_link(request, 'Daily Absent', 'core:daily_absent', 'fa-file-excel', 'daily_absent'),
            _make_link(request, 'Attendance Sheet', 'core:attendance_sheet_manager', 'fa-table', 'attendance_sheet_manager'),
            _make_link(request, 'Subject-wise Attendance', 'core:attendance_sheet_subjectwise_manager', 'fa-book-open', 'attendance_sheet_subjectwise_manager'),
            _make_link(request, 'Lecture Adjustment', 'core:lecture_adjustment', 'fa-exchange-alt', 'lecture_adjustment'),
            _make_link(request, 'Compile Attendance', 'core:compile_attendance', 'fa-file-alt', 'compile_attendance'),
            _make_link(request, 'Overall Attendance', 'core:overall_attendance', 'fa-file-excel', 'overall_attendance'),
            _make_link(request, 'Batchwise Attendance', 'core:admin_batchwise_attendance_manager', 'fa-file-excel', 'admin_batchwise_attendance_manager'),
            _make_link(request, 'Notifications', 'core:admin_notifications', 'fa-bell', 'admin_notifications'),
            _make_link(
                request,
                'Performance Students',
                'core:admin_performance_students',
                'fa-chart-pie',
                'admin_performance_students',
                active_name_extra=(
                    'admin_student_analytics',
                    'admin_mark_analytics',
                    'admin_detailed_mark_analytics',
                    'admin_marks_report',
                    'exam_phases_list',
                    'exam_phase_detail',
                    'exam_phase_upload_marks',
                ),
            ),
        ]
        if is_super_admin:
            links.insert(2, _make_link(request, 'Departmental admins', 'core:departmental_admin_list', 'fa-user-shield', 'departmental_admin_list'))
            links.insert(3, _make_link(request, 'Departmental HODs', 'core:departmental_hod_list', 'fa-user-tie', 'departmental_hod_list'))
            try:
                sa_lock_idx = next(i for i, l in enumerate(links) if l.get('active_name') == 'attendance_lock_setting')
            except StopIteration:
                sa_lock_idx = 12
            links.insert(
                sa_lock_idx + 1,
                _make_link(request, 'Lock Admin Access', 'core:hod_lock_admin_weeks', 'fa-user-lock', 'hod_lock_admin_weeks'),
            )
        elif role == 'hod':
            lock_idx = next((i for i, l in enumerate(links) if l.get('active_name') == 'attendance_lock_setting'), 12)
            links.insert(lock_idx + 1, _make_link(request, 'Lock Admin Access', 'core:hod_lock_admin_weeks', 'fa-user-lock', 'hod_lock_admin_weeks'))
            links.insert(lock_idx + 2, _make_link(request, 'Doubt requests', 'core:hod_doubt_requests', 'fa-comments', 'hod_doubt_requests'))
        if role == 'hod' or is_super_admin:
            try:
                ci = next(i for i, l in enumerate(links) if l.get('active_name') == 'compile_attendance')
            except StopIteration:
                ci = 0
            links.insert(
                ci + 1,
                _make_link(request, 'Daily Report', 'core:daily_report_export', 'fa-file-download', 'daily_report_export'),
            )
            links.insert(
                ci + 2,
                _make_link(request, 'DR weekly load', 'core:admin_faculty_teaching_ds_load', 'fa-balance-scale', 'admin_faculty_teaching_ds_load'),
            )
            try:
                pi = next(i for i, l in enumerate(links) if l.get('active_name') == 'admin_performance_students')
            except StopIteration:
                pi = len(links) - 1
            links.insert(
                pi + 1,
                _make_link(request, 'Management', 'core:admin_faculty_portal_management', 'fa-sliders-h', 'admin_faculty_portal_management'),
            )
        return {'sidebar_links': links, 'is_super_admin': is_super_admin, 'faculty_portal_flags': None}
    if role == 'faculty':
        from core.views import get_faculty_user
        faculty = get_faculty_user(request)
        dept = faculty.department if faculty else None
        flags = {
            'show_doubt': getattr(dept, 'faculty_show_doubt_solving', True) if dept else True,
            'show_dr_load': getattr(dept, 'faculty_show_dr_weekly_load', True) if dept else True,
            'show_mark_analytics': getattr(dept, 'faculty_show_mark_analytics', True) if dept else True,
            'show_marks_report': getattr(dept, 'faculty_show_marks_report', True) if dept else True,
            'show_student_marksheet': getattr(dept, 'faculty_show_student_marksheet', True) if dept else True,
        }
        links = [
            _make_link(request, 'Dashboard', 'core:faculty_dashboard', 'fa-tachometer-alt', 'faculty_dashboard'),
            _make_link(request, 'Mark Attendance', 'core:faculty_attendance_entry', 'fa-edit', 'faculty_attendance_entry'),
        ]
        if flags['show_doubt']:
            links.append(_make_link(request, 'Doubt solving', 'core:faculty_doubt_solving', 'fa-comments', 'faculty_doubt_solving'))
        if flags['show_dr_load']:
            links.append(_make_link(request, 'My DR weekly load', 'core:faculty_dr_load', 'fa-chart-bar', 'faculty_dr_load'))
        links.append(_make_link(request, 'Mentorship Students', 'core:faculty_mentorship', 'fa-users', 'faculty_mentorship'))
        links.append(_make_link(request, 'Student Analytics', 'core:faculty_student_analytics', 'fa-chart-bar', 'faculty_student_analytics'))
        if flags['show_mark_analytics']:
            links.append(_make_link(request, 'Mark Analytics', 'core:faculty_mark_analytics', 'fa-chart-line', 'faculty_mark_analytics'))
        if flags['show_marks_report']:
            links.append(_make_link(request, 'Marks Report', 'core:faculty_marks_report', 'fa-file-alt', 'faculty_marks_report'))
        if flags['show_student_marksheet']:
            links.append(_make_link(request, 'Student Marksheet', 'core:faculty_student_marksheet', 'fa-award', 'faculty_student_marksheet'))
        return {'sidebar_links': links, 'is_super_admin': False, 'faculty_portal_flags': flags}
    if role == 'student':
        links = [
            _make_link(request, 'Dashboard', 'core:student_dashboard', 'fa-tachometer-alt', 'student_dashboard'),
            _make_link(request, 'Attendance Analytics', 'core:student_attendance_analytics', 'fa-chart-line', 'student_attendance_analytics'),
        ]
        return {'sidebar_links': links, 'is_super_admin': False, 'faculty_portal_flags': None}
    return {'sidebar_links': links, 'is_super_admin': is_super_admin, 'faculty_portal_flags': None}
