from django.contrib import admin
from .models import (
    Department, Batch, Subject, Faculty, Student,
    ScheduleSlot, TermPhase, PhaseHoliday, FacultyAttendance, LectureCancellation, ExtraLecture,
    ExamPhase, ExamPhaseSubject, StudentMark, HODWeekLock, FacultyDoubtSession,
    FacultyDoubtRequest, FacultyDoubtRequestStudent,
)

@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    list_display = ('name', 'semester')

@admin.register(Batch)
class BatchAdmin(admin.ModelAdmin):
    list_display = ('name', 'department')
    list_filter = ('department',)

@admin.register(Subject)
class SubjectAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'department')
    list_filter = ('department',)

@admin.register(Faculty)
class FacultyAdmin(admin.ModelAdmin):
    list_display = ('full_name', 'short_name', 'department', 'user')
    list_filter = ('department',)

@admin.register(Student)
class StudentAdmin(admin.ModelAdmin):
    list_display = ('roll_no', 'name', 'batch', 'department')
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
    list_display = ('faculty', 'date', 'batches_list', 'status', 'start_time', 'end_time', 'created_at')
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
