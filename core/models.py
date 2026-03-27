"""
Core models for LJIET_Attendance.
"""
from datetime import date
from django.db import models
from django.contrib.auth.models import User


class Department(models.Model):
    name = models.CharField(max_length=200)
    semester = models.CharField(
        max_length=20,
        blank=True,
        help_text='Semester label (e.g. Roman IV). Used on Daily Report exports.',
    )
    include_extra_lectures_in_attendance = models.BooleanField(
        default=False,
        help_text='If on, extra lectures count in held/percentages like regular slots. If off, faculty can still mark attendance but totals exclude extras.',
    )
    faculty_show_doubt_solving = models.BooleanField(
        default=True,
        help_text='If off, faculty sidebar hides Doubt solving and the feature is blocked.',
    )
    faculty_show_dr_weekly_load = models.BooleanField(
        default=True,
        help_text='If off, faculty sidebar hides My DR weekly load.',
    )
    faculty_show_mark_analytics = models.BooleanField(
        default=True,
        help_text='If off, faculty sidebar hides Mark analytics and related exports.',
    )
    faculty_show_marks_report = models.BooleanField(
        default=True,
        help_text='If off, faculty sidebar hides Marks report.',
    )
    faculty_show_student_marksheet = models.BooleanField(
        default=True,
        help_text='If off, faculty sidebar hides Student marksheet.',
    )

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
    student_phone_number = models.CharField(max_length=20, blank=True)
    parents_contact_number = models.CharField(max_length=20, blank=True)
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


class FacultyCombineDrCache(models.Model):
    """Denormalized DR weekly-combine unit: one row when attendance is saved and matches DR faculty+subject for that slot."""
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='combine_dr_cache_rows')
    faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE, related_name='combine_dr_cache_rows')
    batch = models.ForeignKey(Batch, on_delete=models.CASCADE)
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE)
    date = models.DateField()
    lecture_slot = models.CharField(max_length=50)
    effective_load = models.FloatField(
        default=0.75,
        help_text='DR effective load (0.75 default; ETL with present < 24 → 0.5).',
    )
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        unique_together = ('faculty', 'date', 'batch', 'lecture_slot')
        indexes = [
            models.Index(fields=['department', 'faculty', 'date']),
        ]
        verbose_name = 'Faculty DR combine cache row'
        verbose_name_plural = 'Faculty DR combine cache rows'

    def __str__(self):
        return f'{self.faculty.short_name} {self.date} {self.batch.name} {self.lecture_slot}'


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
    """Per department: after this time (IST) each day, faculty of this department cannot edit attendance. Manual attendance is not time-locked."""
    department = models.OneToOneField(
        Department,
        on_delete=models.CASCADE,
        related_name='attendance_lock_setting',
    )
    lock_hour = models.PositiveSmallIntegerField(default=17)  # 0-23, default 5 PM
    lock_minute = models.PositiveSmallIntegerField(default=0)  # 0-59
    enabled = models.BooleanField(default=False)

    class Meta:
        verbose_name = 'Attendance lock time'
        verbose_name_plural = 'Attendance lock times'

    def __str__(self):
        try:
            tag = self.department.name
        except Exception:
            tag = '—'
        if not self.enabled:
            return f'{tag}: lock disabled'
        return f'{tag}: {self.lock_hour:02d}:{self.lock_minute:02d} IST'


class HODWeekLock(models.Model):
    """Per-department week lock (set by HOD or super admin for the selected dept). When locked, departmental admins cannot edit manual attendance for dates in that week; faculty time lock is unchanged."""
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    phase = models.CharField(max_length=10)  # T1, T2, T3, T4
    week_index = models.PositiveSmallIntegerField()  # 0-based week index within phase
    locked_at = models.DateTimeField(auto_now=True)

    class Meta:
        unique_together = ('department', 'phase', 'week_index')
        verbose_name = 'HOD week lock'
        verbose_name_plural = 'HOD week locks'
        ordering = ['department', 'phase', 'week_index']

    def __str__(self):
        return f'{self.department.name} {self.phase} Week {self.week_index + 1}'


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


class ExtraLecture(models.Model):
    """Extra lecture added on a specific date: batch, time_slot, subject, faculty, room. Reflects everywhere."""
    date = models.DateField()
    batch = models.ForeignKey(Batch, on_delete=models.CASCADE)
    time_slot = models.CharField(max_length=50)
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE)
    faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE)
    room_number = models.CharField(max_length=50, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-date', 'batch', 'time_slot']
        unique_together = ('date', 'batch', 'time_slot')
        verbose_name = 'Extra lecture'
        verbose_name_plural = 'Extra lectures'

    def __str__(self):
        return f"{self.date} {self.batch.name} {self.time_slot} — {self.subject.name} ({self.faculty.short_name})"


class FacultyDoubtSession(models.Model):
    """Legacy: direct logged session (pre–request workflow). Prefer FacultyDoubtRequest for new data."""
    faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE, related_name='doubt_sessions')
    date = models.DateField()
    batch = models.ForeignKey(Batch, on_delete=models.CASCADE)
    student = models.ForeignKey(Student, on_delete=models.CASCADE)
    start_time = models.TimeField()
    end_time = models.TimeField()
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-date', '-start_time', '-pk']
        verbose_name = 'Doubt solving session'
        verbose_name_plural = 'Doubt solving sessions'

    def duration_minutes(self):
        from datetime import datetime, timedelta
        d = self.date
        a = datetime.combine(d, self.start_time)
        b = datetime.combine(d, self.end_time)
        if b <= a:
            b += timedelta(days=1)
        return (b - a).total_seconds() / 60.0

    def duration_hours(self):
        return round(self.duration_minutes() / 60.0, 2)

    def __str__(self):
        return f"{self.faculty.short_name} {self.date} {self.student.roll_no}"


class FacultyDoubtRequest(models.Model):
    """Faculty submits multi-student doubt session; HOD must accept before it counts as DS hours."""
    STATUS_PENDING = 'pending'
    STATUS_ACCEPTED = 'accepted'
    STATUS_REJECTED = 'rejected'
    STATUS_CHOICES = [
        (STATUS_PENDING, 'Pending'),
        (STATUS_ACCEPTED, 'Accepted'),
        (STATUS_REJECTED, 'Rejected'),
    ]

    faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE, related_name='doubt_requests')
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='doubt_requests')
    batch = models.ForeignKey(
        Batch, on_delete=models.CASCADE, null=True, blank=True,
        help_text='Deprecated: use batches. Kept for older rows.',
    )
    batches = models.ManyToManyField(Batch, related_name='faculty_doubt_requests', blank=True)
    date = models.DateField()
    start_time = models.TimeField()
    end_time = models.TimeField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default=STATUS_PENDING)
    created_at = models.DateTimeField(auto_now_add=True)
    reviewed_at = models.DateTimeField(null=True, blank=True)
    reviewed_by = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True, related_name='doubt_requests_reviewed'
    )
    review_notes = models.CharField(max_length=500, blank=True)

    class Meta:
        ordering = ['-date', '-start_time', '-pk']
        verbose_name = 'Doubt solving request'
        verbose_name_plural = 'Doubt solving requests'

    def duration_minutes(self):
        from datetime import datetime, timedelta
        d = self.date
        a = datetime.combine(d, self.start_time)
        b = datetime.combine(d, self.end_time)
        if b <= a:
            b += timedelta(days=1)
        return (b - a).total_seconds() / 60.0

    def nominal_ds_hours(self):
        """Effective DS hours: 1 clock hour = 0.5 effective h (e.g. 60 min clock → 0.5 h)."""
        return self.duration_minutes() / 120.0

    def duration_hours(self):
        """Rounded effective DS hours for display (use nominal_ds_hours() in sums)."""
        return round(self.nominal_ds_hours(), 2)

    def batches_label(self):
        names = sorted(self.batches.values_list('name', flat=True))
        if names:
            return ', '.join(names)
        if self.batch_id:
            return self.batch.name
        return '—'

    def __str__(self):
        return f"{self.faculty.short_name} {self.date} ({self.status})"


class FacultyDoubtRequestStudent(models.Model):
    request = models.ForeignKey(
        FacultyDoubtRequest, on_delete=models.CASCADE, related_name='student_lines'
    )
    student = models.ForeignKey(Student, on_delete=models.CASCADE)

    class Meta:
        unique_together = ('request', 'student')
        ordering = ['student__roll_no']
        verbose_name = 'Doubt request student'
        verbose_name_plural = 'Doubt request students'


class ExamPhase(models.Model):
    """Exam phase (T1, T2, T3, SEE) per department."""
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    name = models.CharField(max_length=50)  # T1, T2, T3, SEE
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        # Logical display order (T1… then SEE) is applied in views/exports via core.exam_phase_order.
        ordering = ['department', 'name']
        unique_together = ('department', 'name')
        verbose_name = 'Exam phase'
        verbose_name_plural = 'Exam phases'

    def __str__(self):
        return f"{self.name} ({self.department.name})"


class ExamPhaseSubject(models.Model):
    """Subjects included in an exam phase."""
    exam_phase = models.ForeignKey(ExamPhase, on_delete=models.CASCADE)
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE)

    class Meta:
        ordering = ['exam_phase', 'subject']
        unique_together = ('exam_phase', 'subject')
        verbose_name = 'Exam phase subject'
        verbose_name_plural = 'Exam phase subjects'

    def __str__(self):
        return f"{self.exam_phase.name} — {self.subject.name}"


class StudentMark(models.Model):
    """Marks for a student in a phase for a subject."""
    student = models.ForeignKey(Student, on_delete=models.CASCADE)
    exam_phase = models.ForeignKey(ExamPhase, on_delete=models.CASCADE)
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE)
    marks_obtained = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['exam_phase', 'subject', 'student']
        unique_together = ('student', 'exam_phase', 'subject')
        verbose_name = 'Student mark'
        verbose_name_plural = 'Student marks'

    def __str__(self):
        return f"{self.student.roll_no} {self.exam_phase.name} {self.subject.name}: {self.marks_obtained}"
