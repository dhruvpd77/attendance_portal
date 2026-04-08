from django.contrib import admin
from .models import (
    InstituteSemester, Department, Batch, Subject, Faculty, FacultyDepartmentMembership, Student,
    ScheduleSlot, TermPhase, PhaseHoliday, FacultyAttendance, LectureCancellation, ExtraLecture,
    ExamPhase, ExamPhaseSubject, StudentMark, HODWeekLock, FacultyDoubtSession,
    FacultyDoubtRequest, FacultyDoubtRequestStudent,
    DepartmentExamProfile, SupervisionExamPhase, SupervisionDuty,
    PaperCheckingPhase, PaperCheckingDuty, PaperCheckingAllocation, PaperCheckingAdjustedShare,
    PaperCheckingSubjectCredit,
    PaperCheckingCompletionRequest,
    PaperSettingPhase, PaperSettingDuty, PaperSettingCompletionRequest,
    DepartmentExamCreditRule,
)

@admin.register(InstituteSemester)
class InstituteSemesterAdmin(admin.ModelAdmin):
    list_display = ('code', 'label', 'sort_order', 'faculty_portal_active', 'created_at')
    list_filter = ('faculty_portal_active',)


@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    list_display = ('name', 'institute_semester', 'faculty_portal_enabled', 'faculty_show_risk_student_info', 'dr_export_semester_label')
    list_filter = ('institute_semester', 'faculty_portal_enabled', 'faculty_show_risk_student_info')

@admin.register(Batch)
class BatchAdmin(admin.ModelAdmin):
    list_display = ('name', 'department')
    list_filter = ('department',)

@admin.register(Subject)
class SubjectAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'department')
    list_filter = ('department',)

@admin.register(FacultyDepartmentMembership)
class FacultyDepartmentMembershipAdmin(admin.ModelAdmin):
    list_display = ('faculty', 'department')
    list_filter = ('department',)
    search_fields = ('faculty__short_name', 'faculty__full_name', 'department__name')


@admin.register(Faculty)
class FacultyAdmin(admin.ModelAdmin):
    list_display = ('full_name', 'short_name', 'department', 'user', 'portal_canonical')
    list_filter = ('department',)
    raw_id_fields = ('portal_canonical',)

@admin.register(Student)
class StudentAdmin(admin.ModelAdmin):
    list_display = ('roll_no', 'name', 'branch', 'batch', 'department')
    list_filter = ('department', 'batch')

@admin.register(ScheduleSlot)
class ScheduleSlotAdmin(admin.ModelAdmin):
    list_display = ('batch', 'day', 'time_slot', 'subject', 'faculty', 'effective_from')
    list_filter = ('department', 'day')

@admin.register(TermPhase)
class TermPhaseAdmin(admin.ModelAdmin):
    list_display = ('department',)

@admin.register(PhaseHoliday)
class PhaseHolidayAdmin(admin.ModelAdmin):
    list_display = ('department', 'phase', 'date')
    list_filter = ('department', 'phase')

@admin.register(FacultyAttendance)
class FacultyAttendanceAdmin(admin.ModelAdmin):
    list_display = ('faculty', 'date', 'batch', 'lecture_slot')
    list_filter = ('date', 'faculty')
    search_fields = ('faculty__short_name', 'batch__name', 'lecture_slot')


@admin.register(LectureCancellation)
class LectureCancellationAdmin(admin.ModelAdmin):
    list_display = ('date', 'batch', 'time_slot')
    list_filter = ('date', 'batch__department')
    search_fields = ('batch__name', 'time_slot')


@admin.register(ExtraLecture)
class ExtraLectureAdmin(admin.ModelAdmin):
    list_display = ('date', 'batch', 'time_slot', 'subject', 'faculty', 'room_number')
    list_filter = ('date', 'batch__department')
    search_fields = ('batch__name', 'time_slot', 'room_number')


@admin.register(FacultyDoubtSession)
class FacultyDoubtSessionAdmin(admin.ModelAdmin):
    list_display = ('faculty', 'date', 'batch', 'student', 'start_time', 'end_time', 'created_at')
    list_filter = ('date', 'batch__department')
    search_fields = ('faculty__short_name', 'student__roll_no', 'student__name')
    raw_id_fields = ('faculty', 'batch', 'student')
    ordering = ('-date', '-start_time')


class FacultyDoubtRequestStudentInline(admin.TabularInline):
    model = FacultyDoubtRequestStudent
    raw_id_fields = ('student',)
    extra = 0


@admin.register(FacultyDoubtRequest)
class FacultyDoubtRequestAdmin(admin.ModelAdmin):
    list_display = ('faculty', 'date', 'batches_list', 'location', 'status', 'start_time', 'end_time', 'created_at')
    list_filter = ('status', 'date', 'department')
    search_fields = ('faculty__short_name', 'batches__name')
    filter_horizontal = ('batches',)
    raw_id_fields = ('faculty', 'department', 'batch', 'reviewed_by')
    inlines = [FacultyDoubtRequestStudentInline]
    ordering = ('-date', '-pk')

    @admin.display(description='Batches')
    def batches_list(self, obj):
        return obj.batches_label()


@admin.register(ExamPhase)
class ExamPhaseAdmin(admin.ModelAdmin):
    list_display = ('name', 'department')
    list_filter = ('department',)


@admin.register(ExamPhaseSubject)
class ExamPhaseSubjectAdmin(admin.ModelAdmin):
    list_display = ('exam_phase', 'subject')
    list_filter = ('exam_phase__department',)


@admin.register(StudentMark)
class StudentMarkAdmin(admin.ModelAdmin):
    list_display = ('student', 'exam_phase', 'subject', 'marks_obtained')
    list_filter = ('exam_phase', 'subject__department')
    search_fields = ('student__roll_no', 'student__name', 'student__enrollment_no')


@admin.register(HODWeekLock)
class HODWeekLockAdmin(admin.ModelAdmin):
    list_display = ('department', 'phase', 'week_index', 'locked_at')
    list_filter = ('department', 'phase')


@admin.register(DepartmentExamProfile)
class DepartmentExamProfileAdmin(admin.ModelAdmin):
    list_display = ('user', 'department', 'parent', 'subunit_code', 'is_hub_coordinator', 'invited_by')
    list_filter = ('department', 'is_hub_coordinator')
    raw_id_fields = ('user', 'parent', 'invited_by')


@admin.register(SupervisionExamPhase)
class SupervisionExamPhaseAdmin(admin.ModelAdmin):
    list_display = ('name', 'department', 'hub_coordinator', 'created_at')
    list_filter = ('department', 'hub_coordinator')


@admin.register(PaperCheckingPhase)
class PaperCheckingPhaseAdmin(admin.ModelAdmin):
    list_display = ('name', 'institute_scope', 'department', 'hub_coordinator', 'created_at')
    list_filter = ('institute_scope', 'department', 'hub_coordinator')


class PaperCheckingAllocationInline(admin.TabularInline):
    model = PaperCheckingAllocation
    extra = 0


@admin.register(PaperCheckingSubjectCredit)
class PaperCheckingSubjectCreditAdmin(admin.ModelAdmin):
    list_display = ('phase', 'subject_name', 'is_practical', 'credit_per_paper_theory', 'credit_online_per_paper', 'credit_offline_per_paper', 'updated_at')
    list_filter = ('phase', 'is_practical')
    search_fields = ('subject_name',)


@admin.register(PaperCheckingAdjustedShare)
class PaperCheckingAdjustedShareAdmin(admin.ModelAdmin):
    list_display = ('duty', 'faculty', 'paper_count', 'created_by', 'created_at')
    list_filter = ('duty__phase',)
    raw_id_fields = ('duty', 'faculty', 'created_by')


@admin.register(PaperCheckingCompletionRequest)
class PaperCheckingCompletionRequestAdmin(admin.ModelAdmin):
    list_display = ('duty', 'faculty', 'status', 'submitted_at', 'decided_at', 'decided_by')
    list_filter = ('status',)
    raw_id_fields = ('duty', 'faculty', 'decided_by')


@admin.register(PaperCheckingDuty)
class PaperCheckingDutyAdmin(admin.ModelAdmin):
    list_display = ('phase', 'faculty', 'exam_date', 'subject_name', 'total_students', 'deadline_date')
    list_filter = ('phase',)
    search_fields = ('subject_name', 'evaluator_short_raw', 'faculty__full_name')
    inlines = [PaperCheckingAllocationInline]
    raw_id_fields = ('phase', 'faculty')


@admin.register(PaperSettingPhase)
class PaperSettingPhaseAdmin(admin.ModelAdmin):
    list_display = ('name', 'institute_scope', 'department', 'hub_coordinator', 'created_at')
    list_filter = ('institute_scope', 'department', 'hub_coordinator')


@admin.register(PaperSettingDuty)
class PaperSettingDutyAdmin(admin.ModelAdmin):
    list_display = ('phase', 'faculty', 'duty_date', 'deadline_date', 'subject_name')
    list_filter = ('phase',)
    search_fields = ('subject_name', 'faculty_short_raw', 'faculty__full_name')
    raw_id_fields = ('phase', 'faculty')


@admin.register(PaperSettingCompletionRequest)
class PaperSettingCompletionRequestAdmin(admin.ModelAdmin):
    list_display = ('duty', 'faculty', 'status', 'submitted_at', 'decided_at', 'decided_by')
    list_filter = ('status',)
    raw_id_fields = ('duty', 'faculty', 'decided_by')


@admin.register(DepartmentExamCreditRule)
class DepartmentExamCreditRuleAdmin(admin.ModelAdmin):
    list_display = ('department', 'task', 'phase_bucket', 'subject_name', 'credit', 'updated_at')
    list_filter = ('task', 'phase_bucket')
    raw_id_fields = ('department',)


@admin.register(SupervisionDuty)
class SupervisionDutyAdmin(admin.ModelAdmin):
    list_display = (
        'phase',
        'faculty',
        'original_faculty',
        'is_proxy',
        'completion_status',
        'supervision_date',
        'division_code',
        'block_no',
        'room_no',
    )
    list_filter = ('phase__department', 'phase', 'completion_status', 'is_proxy')
    search_fields = ('faculty__full_name', 'faculty_name_raw', 'subject_name')
    raw_id_fields = ('phase', 'faculty', 'original_faculty')
