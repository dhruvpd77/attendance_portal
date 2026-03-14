"""
Core models for LJIET_Attendance.
"""
from datetime import date
from django.db import models
from django.contrib.auth.models import User


class Department(models.Model):
    name = models.CharField(max_length=200)
    code = models.CharField(max_length=20, blank=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name


class Batch(models.Model):
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    name = models.CharField(max_length=100)  # e.g. A1, B2, CE-A

    class Meta:
        ordering = ['name']
        unique_together = ('department', 'name')

    def __str__(self):
        return f"{self.name} ({self.department.name})"


class Subject(models.Model):
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    name = models.CharField(max_length=150)
    code = models.CharField(max_length=30, blank=True)

    class Meta:
        ordering = ['name']
        unique_together = ('department', 'name')

    def __str__(self):
        return self.name


class Faculty(models.Model):
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    full_name = models.CharField(max_length=200)
    short_name = models.CharField(max_length=30)  # e.g. UMS, PSK
    email = models.EmailField(blank=True)
    user = models.OneToOneField(
        User, on_delete=models.SET_NULL, null=True, blank=True, related_name='faculty_profile'
    )

    class Meta:
        ordering = ['full_name']
        verbose_name_plural = 'Faculties'

    def __str__(self):
        return f"{self.full_name} ({self.short_name})"


class Student(models.Model):
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    batch = models.ForeignKey(Batch, on_delete=models.CASCADE)
    roll_no = models.CharField(max_length=30)
    enrollment_no = models.CharField(max_length=50, blank=True)
    name = models.CharField(max_length=200)
    email = models.EmailField(blank=True)
    mentor = models.ForeignKey(
        Faculty, on_delete=models.SET_NULL, null=True, blank=True,
        related_name='mentorship_students'
    )
    user = models.OneToOneField(
        User, on_delete=models.SET_NULL, null=True, blank=True, related_name='student_profile'
    )

    class Meta:
        ordering = ['batch', 'roll_no']
        unique_together = ('department', 'batch', 'roll_no')

    def __str__(self):
        return f"{self.roll_no} - {self.name}"


class ScheduleSlot(models.Model):
    """One lecture slot: faculty teaches subject to batch on a weekday at a time.
    effective_from: when this schedule version applies. Past dates use the version valid on that date."""
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE)
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE)
    batch = models.ForeignKey(Batch, on_delete=models.CASCADE)
    day = models.CharField(max_length=20)  # Monday, Tuesday, ...
    time_slot = models.CharField(max_length=50)  # e.g. 08:45-09:45, Lec 1
    effective_from = models.DateField(
        default=date(2000, 1, 1),
        help_text='Date from which this slot applies. Original schedule uses 2000-01-01.'
    )

    class Meta:
        ordering = ['day', 'time_slot']
        unique_together = ('department', 'batch', 'day', 'time_slot', 'effective_from')

    def __str__(self):
        return f"{self.batch.name} {self.day} {self.time_slot} - {self.subject.name} ({self.faculty.short_name})"


class TermPhase(models.Model):
    """T1, T2, T3, T4 date ranges per department."""
    department = models.OneToOneField(Department, on_delete=models.CASCADE)
    t1_start = models.DateField(null=True, blank=True)
    t1_end = models.DateField(null=True, blank=True)
    t2_start = models.DateField(null=True, blank=True)
    t2_end = models.DateField(null=True, blank=True)
    t3_start = models.DateField(null=True, blank=True)
    t3_end = models.DateField(null=True, blank=True)
    t4_start = models.DateField(null=True, blank=True)
    t4_end = models.DateField(null=True, blank=True)

    def __str__(self):
        return f"Term phases for {self.department.name}"


class PhaseHoliday(models.Model):
    """Holiday dates within a term phase. Excluded from lecture days, attendance, and all reports."""
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    phase = models.CharField(max_length=10)  # T1, T2, T3, T4
    date = models.DateField()

    class Meta:
        ordering = ['phase', 'date']
        unique_together = ('department', 'phase', 'date')

    def __str__(self):
        return f"{self.department.name} {self.phase} {self.date}"


class FacultyAttendance(models.Model):
    """One record per faculty per date per batch per lecture slot - absent roll numbers."""
    faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE)
    date = models.DateField()
    batch = models.ForeignKey(Batch, on_delete=models.CASCADE)
    lecture_slot = models.CharField(max_length=50)  # same as ScheduleSlot.time_slot or "Lec 1"
    absent_roll_numbers = models.TextField(blank=True)  # comma-separated
    absent_reasons = models.TextField(blank=True)  # JSON: {"176":"washroom","177":"general"} default general
    remarks = models.CharField(max_length=200, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-date', 'batch', 'lecture_slot']
        unique_together = ('faculty', 'date', 'batch', 'lecture_slot')

    def __str__(self):
        return f"{self.faculty.short_name} {self.date} {self.batch.name} {self.lecture_slot}"


class AttendanceNotificationLog(models.Model):
    """Track when low-attendance emails were sent to students to avoid spam."""
    student = models.ForeignKey(Student, on_delete=models.CASCADE)
    notification_type = models.CharField(max_length=30, default='low_attendance')
    sent_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-sent_at']

    def __str__(self):
        return f"{self.student.roll_no} {self.notification_type} @ {self.sent_at}"


class AttendanceLockSetting(models.Model):
    """Global setting: after this time (IST) each day, faculty cannot edit attendance. Admin manual attendance is never locked."""
    lock_hour = models.PositiveSmallIntegerField(default=17)  # 0-23, default 5 PM
    lock_minute = models.PositiveSmallIntegerField(default=0)  # 0-59
    enabled = models.BooleanField(default=False)

    class Meta:
        verbose_name = 'Attendance lock time'
        verbose_name_plural = 'Attendance lock time'

    def __str__(self):
        if not self.enabled:
            return 'Lock disabled'
        return f'{self.lock_hour:02d}:{self.lock_minute:02d} IST (locked after this time each day)'


class LectureCancellation(models.Model):
    """Lecture cancelled on this date. Excluded from lecture counts and all attendance records."""
    date = models.DateField()
    batch = models.ForeignKey(Batch, on_delete=models.CASCADE)
    time_slot = models.CharField(max_length=50)  # same as ScheduleSlot.time_slot

    class Meta:
        ordering = ['-date', 'batch', 'time_slot']
        unique_together = ('date', 'batch', 'time_slot')
        verbose_name = 'Lecture cancellation'
        verbose_name_plural = 'Lecture cancellations'

    def __str__(self):
        return f"{self.date} {self.batch.name} {self.time_slot} (cancelled)"


class LectureAdjustment(models.Model):
    """One-off override: on this date, this batch/slot was taken by new_faculty for new_subject (e.g. substitute)."""
    date = models.DateField()
    batch = models.ForeignKey(Batch, on_delete=models.CASCADE)
    time_slot = models.CharField(max_length=50)  # same as ScheduleSlot.time_slot
    original_faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE, related_name='adjustments_original')
    original_subject = models.ForeignKey(Subject, on_delete=models.CASCADE, related_name='adjustments_original')
    new_faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE, related_name='adjustments_new')
    new_subject = models.ForeignKey(Subject, on_delete=models.CASCADE, related_name='adjustments_new')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-date', 'batch', 'time_slot']
        unique_together = ('date', 'batch', 'time_slot')

    def __str__(self):
        return f"{self.date} {self.batch.name} {self.time_slot} → {self.new_subject.name} ({self.new_faculty.short_name})"
