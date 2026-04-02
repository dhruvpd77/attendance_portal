"""Faculty exam portal: History vs current is driven by InstituteSemester.faculty_portal_active.

When super admin turns Faculty portal **off** for an academic semester, attendance hides as before,
but exam duties & credits for that semester appear under **History** (read-only). Pending vs approved
states are unchanged so coordinators can still see real pending/completed in History.
"""
from __future__ import annotations

from core.models import Faculty, InstituteSemester


def semester_for_supervision_duty(duty) -> InstituteSemester | None:
    return getattr(duty.phase, 'institute_semester', None)


def semester_for_paper_phase(phase, faculty: Faculty | None) -> InstituteSemester | None:
    if phase and getattr(phase, 'institute_semester_id', None):
        return phase.institute_semester
    if faculty and faculty.department_id:
        return faculty.department.institute_semester
    return None


def is_semester_faculty_portal_closed(sem: InstituteSemester | None) -> bool:
    """True when exam rows for this semester belong in History (read-only for faculty)."""
    if sem is None:
        return False
    return not sem.faculty_portal_active


def supervision_duty_in_faculty_exam_history(duty) -> bool:
    return is_semester_faculty_portal_closed(semester_for_supervision_duty(duty))


def paper_checking_duty_in_faculty_exam_history(duty, faculty: Faculty | None) -> bool:
    return is_semester_faculty_portal_closed(semester_for_paper_phase(duty.phase, faculty))


def paper_setting_duty_in_faculty_exam_history(duty, faculty: Faculty | None) -> bool:
    return is_semester_faculty_portal_closed(semester_for_paper_phase(duty.phase, faculty))


def paper_checking_completion_in_faculty_exam_history(req, faculty: Faculty | None) -> bool:
    return paper_checking_duty_in_faculty_exam_history(req.duty, faculty)


def paper_setting_completion_in_faculty_exam_history(req, faculty: Faculty | None) -> bool:
    return paper_setting_duty_in_faculty_exam_history(req.duty, faculty)


def faculty_has_exam_portal_history_rows(faculty: Faculty | None) -> bool:
    """True if faculty has any exam row tied to a closed (faculty-portal-off) academic semester (sidebar History link)."""
    from django.db.models import Q

    from core.models import PaperCheckingDuty, PaperSettingDuty, SupervisionDuty

    if not faculty:
        return False
    closed = Q(phase__institute_semester__faculty_portal_active=False)
    if SupervisionDuty.objects.filter(faculty=faculty).filter(closed).exists():
        return True
    if (
        PaperCheckingDuty.objects.filter(Q(faculty=faculty) | Q(adjusted_shares__faculty=faculty))
        .filter(closed)
        .exists()
    ):
        return True
    if PaperSettingDuty.objects.filter(faculty=faculty).filter(closed).exists():
        return True
    return False
