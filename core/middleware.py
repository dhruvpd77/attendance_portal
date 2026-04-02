from django.shortcuts import redirect

from core.semester_scope import (
    child_must_select_exam_context,
    coordinator_must_select_exam_context,
    get_active_institute_semester,
)


class FacultyWorkingDepartmentMiddleware:
    """Attach request.faculty_working_department for faculty role (one query per request)."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        request.faculty_working_department = None
        if request.user.is_authenticated:
            try:
                if getattr(request.user, 'role_profile', None) and request.user.role_profile.role == 'faculty':
                    from core.faculty_scope import get_faculty_working_department

                    request.faculty_working_department = get_faculty_working_department(request)
            except Exception:
                pass
        return self.get_response(request)


class InstituteSemesterMiddleware:
    """Attach request.active_institute_semester for templates and views."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        request.active_institute_semester = get_active_institute_semester(request)
        return self.get_response(request)


class CoordinatorExamContextMiddleware:
    """Force institute / sub-unit coordinators to pick academic-semester context when they have several."""

    _skip_suffixes = (
        '/exam-dept/select-context/',
        '/exam-dept/select-child-context/',
    )

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        if request.user.is_authenticated:
            path = request.path
            if path.startswith('/portal/exam-dept/') and not any(
                path.endswith(s) for s in self._skip_suffixes
            ):
                try:
                    role = request.user.role_profile.role
                except Exception:
                    role = None
                if role == 'dept_exam_parent' and coordinator_must_select_exam_context(request):
                    return redirect('core:dept_exam_select_context')
                if role == 'dept_exam_child' and child_must_select_exam_context(request):
                    return redirect('core:dept_exam_child_select_context')
        return self.get_response(request)

