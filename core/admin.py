from django.contrib import admin
from .models import (
    Department, Batch, Subject, Faculty, Student,
    ScheduleSlot, TermPhase, PhaseHoliday, FacultyAttendance, LectureCancellation, ExtraLecture,
)

@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    list_display = ('name', 'code')

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
