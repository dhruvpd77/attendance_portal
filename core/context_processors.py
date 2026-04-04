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
    empty = {
        'sidebar_links': links,
        'is_super_admin': False,
        'faculty_portal_flags': None,
        'faculty_working_department': None,
        'faculty_department_choices': [],
    }
    if not request.user.is_authenticated:
        return {
            'sidebar_links': [],
            'is_super_admin': False,
            'faculty_portal_flags': None,
            'is_exam_admin': False,
            'is_exam_coordination': False,
            'faculty_working_department': None,
            'faculty_department_choices': [],
        }
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
            _make_link(
                request,
                'Attendance Sheet',
                'core:attendance_sheet_manager',
                'fa-table',
                'attendance_sheet_manager',
                active_name_extra=('attendance_sheet_subjectwise_manager',),
            ),
            _make_link(request, 'Lecture Adjustment', 'core:lecture_adjustment', 'fa-exchange-alt', 'lecture_adjustment'),
            _make_link(request, 'Compile Attendance', 'core:compile_attendance', 'fa-file-alt', 'compile_attendance'),
            _make_link(request, 'Overall Attendance', 'core:overall_attendance', 'fa-file-excel', 'overall_attendance'),
            _make_link(request, 'Batchwise Attendance', 'core:admin_batchwise_attendance_manager', 'fa-file-excel', 'admin_batchwise_attendance_manager'),
            _make_link(request, 'Notifications', 'core:admin_notifications', 'fa-bell', 'admin_notifications'),
            _make_link(
                request,
                'Risk students (Excel)',
                'core:admin_risk_students_export',
                'fa-file-excel',
                'admin_risk_students_export',
                active_name_extra=('admin_risk_students_excel',),
            ),
            _make_link(
                request,
                'Risk student info (logs)',
                'core:admin_risk_student_info',
                'fa-phone',
                'admin_risk_student_info',
                active_name_extra=('admin_risk_student_info_save', 'admin_risk_student_info_excel'),
            ),
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
            links.insert(
                4,
                _make_link(
                    request,
                    'Exam Management',
                    'core:admin_exam_management',
                    'fa-layer-group',
                    'admin_exam_management',
                    active_name_extra=(
                        'exam_section_dashboard',
                        'exam_section_working_semesters',
                        'exam_section_create_coordinator',
                        'exam_section_create_operator',
                        'exam_section_delete_coordinator',
                        'exam_section_delete_operator',
                        'exam_section_edit_coordinator',
                        'exam_section_edit_operator',
                        'dept_exam_dr_report_excel',
                        'exam_section_daily_dr_excel',
                        'exam_section_credit_analytics',
                        'exam_section_credit_analytics_excel',
                        'exam_section_supervision_credit_analytics',
                        'exam_section_supervision_credit_analytics_excel',
                        'dr_facilities_dashboard',
                        'dr_facilities_export_department',
                        'dr_facilities_export_faculty',
                        'paper_checking_dashboard',
                        'paper_checking_phase_detail',
                        'paper_checking_phase_upload',
                        'paper_checking_phase_clear',
                        'paper_checking_child_approve_completion',
                        'paper_checking_child_dismiss_completion',
                        'paper_checking_child_save_adjustment',
                        'paper_checking_phase_subject_credits_save',
                        'paper_setting_dashboard',
                        'paper_setting_phase_detail',
                        'paper_setting_phase_upload',
                        'paper_setting_phase_clear',
                        'paper_setting_completion_approve',
                        'paper_setting_completion_dismiss',
                        'exam_credit_settings',
                        'section_credit_report',
                        'section_credit_report_excel',
                        'overall_faculty_report',
                        'overall_faculty_report_excel',
                        'paper_checking_phase_rename',
                        'paper_checking_phase_delete',
                        'paper_setting_phase_rename',
                        'paper_setting_phase_delete',
                    ),
                ),
            )
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
        from core.views import get_admin_department

        _risk_admin_nav = frozenset(
            {
                'admin_risk_students_export',
                'admin_risk_students_excel',
                'admin_risk_student_info',
                'admin_risk_student_info_save',
                'admin_risk_student_info_excel',
            }
        )
        _admin_dept_nav = get_admin_department(request)
        if _admin_dept_nav and not getattr(_admin_dept_nav, 'faculty_show_risk_student_info', True):

            def _keep_admin_nav_item(lk):
                if lk.get('active_name') in _risk_admin_nav:
                    return False
                return not any(x in _risk_admin_nav for x in (lk.get('active_name_extra') or []))

            links = [lk for lk in links if _keep_admin_nav_item(lk)]
        return {
            'sidebar_links': links,
            'is_super_admin': is_super_admin,
            'faculty_portal_flags': None,
            'is_exam_admin': False,
            'is_exam_coordination': False,
            'faculty_working_department': None,
            'faculty_department_choices': [],
        }
    section_report_extra = (
        'section_credit_report',
        'section_credit_report_excel',
        'overall_faculty_report',
        'overall_faculty_report_excel',
    )
    if role == 'exam_section':
        from core.models import InstituteSemester
        from core.semester_scope import exam_section_working_semester_ids, is_exam_section_operator

        exam_section_working_semesters_nav: list[dict] = []
        if is_exam_section_operator(request):
            wids = exam_section_working_semester_ids(request)
            if wids:
                exam_section_working_semesters_nav = list(
                    InstituteSemester.objects.filter(pk__in=wids)
                    .order_by('-sort_order', '-pk')
                    .values('pk', 'label')
                )

        dr_extra = (
            'dr_facilities_export_department',
            'dr_facilities_export_faculty',
        )
        _pex = (
            'paper_checking_phase_detail',
            'paper_checking_phase_upload',
            'paper_checking_phase_clear',
            'paper_checking_child_approve_completion',
            'paper_checking_child_dismiss_completion',
            'paper_checking_child_save_adjustment',
            'paper_checking_phase_subject_credits_save',
            'paper_checking_phase_rename',
            'paper_checking_phase_delete',
        )
        _pse = (
            'paper_setting_phase_detail',
            'paper_setting_phase_upload',
            'paper_setting_phase_clear',
            'paper_setting_completion_approve',
            'paper_setting_completion_dismiss',
            'exam_credit_settings',
            'paper_setting_phase_rename',
            'paper_setting_phase_delete',
        )
        links = [
            _make_link(
                request,
                'Working academic semesters',
                'core:exam_section_working_semesters',
                'fa-calendar-week',
                'exam_section_working_semesters',
            ),
            _make_link(
                request,
                'Exam section',
                'core:exam_section_dashboard',
                'fa-building-columns',
                'exam_section_dashboard',
                active_name_extra=(
                    'exam_section_create_coordinator',
                    'exam_section_create_operator',
                    'exam_section_delete_coordinator',
                    'exam_section_delete_operator',
                    'exam_section_edit_coordinator',
                    'exam_section_edit_operator',
                    'dept_exam_dr_report_excel',
                    'exam_section_daily_dr_excel',
                    'exam_section_credit_analytics',
                    'exam_section_credit_analytics_excel',
                    'exam_section_supervision_credit_analytics',
                    'exam_section_supervision_credit_analytics_excel',
                    'dr_facilities_dashboard',
                    'paper_checking_dashboard',
                    'paper_setting_dashboard',
                    'paper_setting_completion_approve',
                    'paper_setting_completion_dismiss',
                    'exam_credit_settings',
                )
                + dr_extra
                + _pex
                + _pse
                + section_report_extra,
            ),
            _make_link(
                request,
                'Paper-check credits',
                'core:exam_section_credit_analytics',
                'fa-chart-line',
                'exam_section_credit_analytics',
                active_name_extra=('exam_section_credit_analytics_excel',),
            ),
            _make_link(
                request,
                'Supervision credits',
                'core:exam_section_supervision_credit_analytics',
                'fa-user-check',
                'exam_section_supervision_credit_analytics',
                active_name_extra=('exam_section_supervision_credit_analytics_excel',),
            ),
            _make_link(
                request,
                'DR facilities',
                'core:dr_facilities_dashboard',
                'fa-table-list',
                'dr_facilities_dashboard',
                active_name_extra=dr_extra,
            ),
            _make_link(
                request,
                'Paper checking',
                'core:paper_checking_dashboard',
                'fa-check-double',
                'paper_checking_dashboard',
                active_name_extra=_pex,
            ),
            _make_link(
                request,
                'Paper setting',
                'core:paper_setting_dashboard',
                'fa-pen-fancy',
                'paper_setting_dashboard',
                active_name_extra=_pse,
            ),
            _make_link(
                request,
                'Credit settings',
                'core:exam_credit_settings',
                'fa-sliders',
                'exam_credit_settings',
            ),
            _make_link(
                request,
                'Section credit report',
                'core:section_credit_report',
                'fa-file-invoice-dollar',
                'section_credit_report',
                active_name_extra=('section_credit_report_excel',),
            ),
            _make_link(
                request,
                'Overall faculty report',
                'core:overall_faculty_report',
                'fa-users-cog',
                'overall_faculty_report',
                active_name_extra=('overall_faculty_report_excel',),
            ),
        ]
        return {
            'sidebar_links': links,
            'is_super_admin': False,
            'faculty_portal_flags': None,
            'is_exam_admin': False,
            'is_exam_coordination': True,
            'faculty_working_department': None,
            'faculty_department_choices': [],
            'exam_section_working_semesters_nav': exam_section_working_semesters_nav,
        }
    if role == 'dept_exam_parent':
        dr_extra = ('dr_facilities_export_department', 'dr_facilities_export_faculty')
        _pex = (
            'paper_checking_phase_detail',
            'paper_checking_phase_upload',
            'paper_checking_phase_clear',
            'paper_checking_child_approve_completion',
            'paper_checking_child_dismiss_completion',
            'paper_checking_child_save_adjustment',
            'paper_checking_phase_subject_credits_save',
            'paper_checking_phase_rename',
            'paper_checking_phase_delete',
        )
        _pse = (
            'paper_setting_phase_detail',
            'paper_setting_phase_upload',
            'paper_setting_phase_clear',
            'paper_setting_completion_approve',
            'paper_setting_completion_dismiss',
            'exam_credit_settings',
            'paper_setting_phase_rename',
            'paper_setting_phase_delete',
        )
        links = [
            _make_link(
                request,
                'Supervision',
                'core:dept_exam_dashboard',
                'fa-clipboard-list',
                'dept_exam_dashboard',
                active_name_extra=(
                    'dept_exam_select_context',
                    'dept_exam_link_department',
                    'dept_exam_phase_detail',
                    'dept_exam_phase_upload',
                    'dept_exam_phase_clear_duties',
                    'dept_exam_phase_create_faculty_assign',
                    'dept_exam_phase_add',
                    'dept_exam_phase_rename',
                    'dept_exam_phase_delete',
                    'dept_exam_child_create',
                    'dept_exam_child_edit',
                    'dept_exam_child_delete',
                    'dept_exam_dr_report_excel',
                    'dept_exam_daily_dr_excel',
                    'dept_exam_credit_analytics',
                    'dept_exam_credit_analytics_excel',
                    'dept_exam_supervision_credit_analytics',
                    'dept_exam_supervision_credit_analytics_excel',
                    'dr_facilities_dashboard',
                    'paper_checking_dashboard',
                    'paper_setting_dashboard',
                    'paper_setting_completion_approve',
                    'paper_setting_completion_dismiss',
                    'exam_credit_settings',
                )
                + dr_extra
                + _pex
                + _pse
                + section_report_extra,
            ),
            _make_link(
                request,
                'Paper-check credits',
                'core:dept_exam_credit_analytics',
                'fa-chart-line',
                'dept_exam_credit_analytics',
                active_name_extra=('dept_exam_credit_analytics_excel',),
            ),
            _make_link(
                request,
                'Supervision credits',
                'core:dept_exam_supervision_credit_analytics',
                'fa-user-check',
                'dept_exam_supervision_credit_analytics',
                active_name_extra=('dept_exam_supervision_credit_analytics_excel',),
            ),
            _make_link(
                request,
                'DR facilities',
                'core:dr_facilities_dashboard',
                'fa-table-list',
                'dr_facilities_dashboard',
                active_name_extra=dr_extra,
            ),
            _make_link(
                request,
                'Paper checking',
                'core:paper_checking_dashboard',
                'fa-check-double',
                'paper_checking_dashboard',
                active_name_extra=_pex,
            ),
            _make_link(
                request,
                'Paper setting',
                'core:paper_setting_dashboard',
                'fa-pen-fancy',
                'paper_setting_dashboard',
                active_name_extra=_pse,
            ),
            _make_link(
                request,
                'Credit settings',
                'core:exam_credit_settings',
                'fa-sliders',
                'exam_credit_settings',
            ),
            _make_link(
                request,
                'Section credit report',
                'core:section_credit_report',
                'fa-file-invoice-dollar',
                'section_credit_report',
                active_name_extra=('section_credit_report_excel',),
            ),
            _make_link(
                request,
                'Overall faculty report',
                'core:overall_faculty_report',
                'fa-users-cog',
                'overall_faculty_report',
                active_name_extra=('overall_faculty_report_excel',),
            ),
        ]
        return {
            'sidebar_links': links,
            'is_super_admin': False,
            'faculty_portal_flags': None,
            'is_exam_admin': False,
            'is_exam_coordination': True,
            'faculty_working_department': None,
            'faculty_department_choices': [],
        }
    if role == 'dept_exam_child':
        dr_extra = ('dr_facilities_export_department', 'dr_facilities_export_faculty')
        _pex = (
            'paper_checking_phase_detail',
            'paper_checking_phase_upload',
            'paper_checking_phase_clear',
            'paper_checking_child_approve_completion',
            'paper_checking_child_dismiss_completion',
            'paper_checking_child_save_adjustment',
            'paper_checking_phase_subject_credits_save',
        )
        _pse = (
            'paper_setting_phase_detail',
            'paper_setting_phase_upload',
            'paper_setting_phase_clear',
            'paper_setting_completion_approve',
            'paper_setting_completion_dismiss',
        )
        links = [
            _make_link(
                request,
                'Sub-unit',
                'core:dept_exam_dashboard',
                'fa-layer-group',
                'dept_exam_dashboard',
                active_name_extra=(
                    'dept_exam_child_select_context',
                    'dept_exam_dr_report_excel',
                    'dept_exam_daily_dr_excel',
                    'dept_exam_credit_analytics',
                    'dept_exam_credit_analytics_excel',
                    'dept_exam_supervision_credit_analytics',
                    'dept_exam_supervision_credit_analytics_excel',
                    'dept_exam_proxy_supervision',
                    'dr_facilities_dashboard',
                    'paper_checking_dashboard',
                    'paper_setting_dashboard',
                    'paper_setting_completion_approve',
                    'paper_setting_completion_dismiss',
                    'section_credit_report',
                    'section_credit_report_excel',
                    'overall_faculty_report',
                    'overall_faculty_report_excel',
                )
                + dr_extra
                + _pex
                + _pse,
            ),
            _make_link(
                request,
                'Paper-check credits',
                'core:dept_exam_credit_analytics',
                'fa-chart-line',
                'dept_exam_credit_analytics',
                active_name_extra=('dept_exam_credit_analytics_excel',),
            ),
            _make_link(
                request,
                'Supervision credits',
                'core:dept_exam_supervision_credit_analytics',
                'fa-user-check',
                'dept_exam_supervision_credit_analytics',
                active_name_extra=('dept_exam_supervision_credit_analytics_excel',),
            ),
            _make_link(
                request,
                'DR facilities',
                'core:dr_facilities_dashboard',
                'fa-table-list',
                'dr_facilities_dashboard',
                active_name_extra=dr_extra,
            ),
            _make_link(
                request,
                'Paper checking',
                'core:paper_checking_dashboard',
                'fa-check-double',
                'paper_checking_dashboard',
                active_name_extra=_pex,
            ),
            _make_link(
                request,
                'Paper setting',
                'core:paper_setting_dashboard',
                'fa-pen-fancy',
                'paper_setting_dashboard',
                active_name_extra=_pse,
            ),
            _make_link(
                request,
                'Section credit report',
                'core:section_credit_report',
                'fa-file-invoice-dollar',
                'section_credit_report',
                active_name_extra=('section_credit_report_excel',),
            ),
            _make_link(
                request,
                'Overall faculty report',
                'core:overall_faculty_report',
                'fa-users-cog',
                'overall_faculty_report',
                active_name_extra=('overall_faculty_report_excel',),
            ),
        ]
        return {
            'sidebar_links': links,
            'is_super_admin': False,
            'faculty_portal_flags': None,
            'is_exam_admin': False,
            'is_exam_coordination': True,
            'faculty_working_department': None,
            'faculty_department_choices': [],
        }
    if role == 'exam_admin':
        links = [
            _make_link(
                request,
                'Exam Analytics Hub',
                'core:exam_admin_dashboard',
                'fa-chart-line',
                'exam_admin_dashboard',
                active_name_extra=(
                    'exam_admin_excel_compiled_dept',
                    'exam_admin_excel_compiled_all',
                    'exam_admin_excel_subject_by_dept',
                    'exam_admin_excel_subject_all',
                    'exam_admin_excel_phase_compile',
                    'exam_admin_excel_risk',
                    'exam_admin_excel_top10',
                ),
            ),
            _make_link(
                request,
                'Exam phases & uploads',
                'core:exam_admin_exam_phases_list',
                'fa-file-upload',
                'exam_admin_exam_phases_list',
                active_name_extra=('exam_admin_exam_phase_detail',),
            ),
            _make_link(request, 'Mark analytics (table)', 'core:exam_admin_mark_analytics', 'fa-table', 'exam_admin_mark_analytics'),
        ]
        return {
            'sidebar_links': links,
            'is_super_admin': False,
            'faculty_portal_flags': None,
            'is_exam_admin': True,
            'is_exam_coordination': False,
            'faculty_working_department': None,
            'faculty_department_choices': [],
        }
    if role == 'faculty':
        from core.exam_faculty_portal_visibility import faculty_has_exam_portal_history_rows
        from core.faculty_scope import (
            faculty_portal_member_departments_qs,
            get_faculty_exam_context_department,
            get_faculty_working_department,
        )
        from core.views import get_faculty_user, user_can_manage_faculty_portal_settings

        faculty = get_faculty_user(request)
        dept = get_faculty_working_department(request) if faculty else None
        exam_dept = get_faculty_exam_context_department(request) if faculty else None
        dept_choices = list(faculty_portal_member_departments_qs(faculty)) if faculty else []
        dept_selected = dept or exam_dept
        flags = {
            'show_doubt': getattr(dept, 'faculty_show_doubt_solving', True) if dept else True,
            'show_dr_load': getattr(dept, 'faculty_show_dr_weekly_load', True) if dept else True,
            'show_mark_analytics': getattr(dept, 'faculty_show_mark_analytics', True) if dept else True,
            'show_marks_report': getattr(dept, 'faculty_show_marks_report', True) if dept else True,
            'show_student_marksheet': getattr(dept, 'faculty_show_student_marksheet', True) if dept else True,
            'show_exam_duties': getattr(exam_dept, 'faculty_show_exam_duties', False) if exam_dept else False,
            'show_exam_credits': getattr(exam_dept, 'faculty_show_exam_credits_analytics', False)
            if exam_dept
            else False,
            'show_student_analytics_nav': getattr(dept, 'faculty_show_student_analytics', False) if dept else False,
            'show_risk_student_info': getattr(dept, 'faculty_show_risk_student_info', True) if dept else True,
            'show_exam_history_nav': faculty_has_exam_portal_history_rows(faculty) if faculty else False,
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
        if flags['show_risk_student_info']:
            links.append(
                _make_link(
                    request,
                    'Risk student info',
                    'core:faculty_risk_student_info',
                    'fa-phone',
                    'faculty_risk_student_info',
                    active_name_extra=('faculty_risk_student_info_excel',),
                )
            )
        if flags['show_exam_duties']:
            links.append(
                _make_link(
                    request,
                    'My exam duties',
                    'core:faculty_supervision_duties',
                    'fa-clipboard-check',
                    'faculty_supervision_duties',
                    active_name_extra=(
                        'faculty_supervision_duty_complete',
                        'faculty_paper_checking_completion_request',
                        'faculty_paper_setting_completion_request',
                    ),
                )
            )
        if flags['show_exam_credits']:
            links.append(
                _make_link(
                    request,
                    'Exam credits & analytics',
                    'core:faculty_exam_credits_analytics',
                    'fa-chart-pie',
                    'faculty_exam_credits_analytics',
                )
            )
        if flags['show_exam_history_nav']:
            if flags['show_exam_duties']:
                links.append(
                    _make_link(
                        request,
                        'History — exam duties',
                        'core:faculty_exam_history_duties',
                        'fa-clock-rotate-left',
                        'faculty_exam_history_duties',
                    )
                )
            if flags['show_exam_credits']:
                links.append(
                    _make_link(
                        request,
                        'History — credits & analytics',
                        'core:faculty_exam_history_credits_analytics',
                        'fa-archive',
                        'faculty_exam_history_credits_analytics',
                    )
                )
        if flags['show_student_analytics_nav']:
            links.append(
                _make_link(request, 'Student Analytics', 'core:faculty_student_analytics', 'fa-chart-bar', 'faculty_student_analytics')
            )
        if flags['show_mark_analytics']:
            links.append(_make_link(request, 'Mark Analytics', 'core:faculty_mark_analytics', 'fa-chart-line', 'faculty_mark_analytics'))
        if flags['show_marks_report']:
            links.append(_make_link(request, 'Marks Report', 'core:faculty_marks_report', 'fa-file-alt', 'faculty_marks_report'))
        if flags['show_student_marksheet']:
            links.append(_make_link(request, 'Student Marksheet', 'core:faculty_student_marksheet', 'fa-award', 'faculty_student_marksheet'))
        if user_can_manage_faculty_portal_settings(request):
            links.append(
                _make_link(
                    request,
                    'Management',
                    'core:admin_faculty_portal_management',
                    'fa-sliders-h',
                    'admin_faculty_portal_management',
                )
            )
        return {
            'sidebar_links': links,
            'is_super_admin': False,
            'faculty_portal_flags': flags,
            'is_exam_admin': False,
            'is_exam_coordination': False,
            'faculty_working_department': dept,
            'faculty_exam_context_department': exam_dept,
            'faculty_department_selected': dept_selected,
            'faculty_department_choices': dept_choices,
        }
    if role == 'student':
        links = [
            _make_link(request, 'Dashboard', 'core:student_dashboard', 'fa-tachometer-alt', 'student_dashboard'),
            _make_link(request, 'Attendance Analytics', 'core:student_attendance_analytics', 'fa-chart-line', 'student_attendance_analytics'),
        ]
        return {
            'sidebar_links': links,
            'is_super_admin': False,
            'faculty_portal_flags': None,
            'is_exam_admin': False,
            'is_exam_coordination': False,
            'faculty_working_department': None,
            'faculty_department_choices': [],
        }
    return {
        'sidebar_links': links,
        'is_super_admin': is_super_admin,
        'faculty_portal_flags': None,
        'is_exam_admin': False,
        'is_exam_coordination': False,
        'faculty_working_department': None,
        'faculty_department_choices': [],
    }
