"""
Core models for LJIET_Attendance.
"""
from datetime import date
from django.db import models
from django.contrib.auth.models import User


class InstituteSemester(models.Model):
    """Academic period for the whole institute (e.g. SY EVEN 2026). All departments, attendance, and exam data belong to one row."""

    code = models.SlugField(
        max_length=64,
        unique=True,
        help_text='Stable id shown in URLs/exports, e.g. SY_EVEN_2026',
    )
    label = models.CharField(max_length=200)
    sort_order = models.PositiveSmallIntegerField(
        default=0,
        help_text='Higher sorts first in pickers.',
    )
    created_at = models.DateTimeField(auto_now_add=True)
    faculty_portal_active = models.BooleanField(
        default=False,
        help_text='When off, faculty users cannot see attendance, mentorship, or other portal data for departments in this period. Super admin and departmental admin views are unchanged. Turn on when the period is ready for faculty. Exam duties and credits for this semester still appear under History (read-only) when exam menu is enabled for the department.',
    )
    risk_attendance_min_percent = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        default=75,
        help_text='Institute default: cumulative attendance below this %% is at-risk (all term phases). Departments can override.',
    )

    class Meta:
        ordering = ['-sort_order', '-pk']
        verbose_name = 'Academic semester'
        verbose_name_plural = 'Academic semesters'

    def __str__(self):
        return self.label


class Department(models.Model):
    name = models.CharField(max_length=200)
    institute_semester = models.ForeignKey(
        InstituteSemester,
        on_delete=models.CASCADE,
        related_name='departments',
    )
    dr_export_semester_label = models.CharField(
        max_length=20,
        blank=True,
        help_text='Short label for Daily Report exports (e.g. Roman IV) — not the institute semester record.',
    )
    include_extra_lectures_in_attendance = models.BooleanField(
        default=False,
        help_text='If on, extra lectures count in held/percentages like regular slots. If off, faculty can still mark attendance but totals exclude extras.',
    )
    faculty_show_mark_attendance = models.BooleanField(
        default=True,
        help_text='If off, faculty sidebar hides Mark Attendance and saving attendance is blocked.',
    )
    faculty_show_mentorship = models.BooleanField(
        default=True,
        help_text='If off, faculty sidebar hides Mentorship Students.',
    )
    risk_attendance_min_percent = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True,
        help_text='If set, overrides institute semester default for at-risk attendance %.',
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
    faculty_show_exam_duties = models.BooleanField(
        default=False,
        help_text='If off, faculty sidebar hides My exam duties and related actions.',
    )
    faculty_show_exam_credits_analytics = models.BooleanField(
        default=False,
        help_text='If off, faculty sidebar hides Exam credits & analytics.',
    )
    faculty_show_student_analytics = models.BooleanField(
        default=False,
        help_text='If off, faculty sidebar hides Student Analytics.',
    )
    faculty_show_risk_student_info = models.BooleanField(
        default=True,
        help_text='If off, Risk student info (faculty + admin), student-wise Excel, and admin Risk students (Excel) export are hidden and blocked.',
    )
    faculty_portal_enabled = models.BooleanField(
        default=False,
        help_text='Master switch: when off, this department is hidden from faculty portal (attendance, mentorship, etc.). HOD / super admin dashboards are unchanged. Enable in Management when this division should use the faculty portal.',
    )

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name


class InstituteExamPhaseRiskThreshold(models.Model):
    """Super-admin defaults: marks below this value are at-risk for the phase (per academic semester)."""

    institute_semester = models.ForeignKey(
        InstituteSemester, on_delete=models.CASCADE, related_name='exam_phase_risk_thresholds'
    )
    phase_name = models.CharField(max_length=50, help_text='Exam phase name, e.g. T1, T2, SEE')
    fail_below_marks = models.DecimalField(
        max_digits=6,
        decimal_places=2,
        default=9,
        help_text='At-risk if marks obtained is strictly less than this (same as legacy &lt; 9).',
    )

    class Meta:
        ordering = ['institute_semester', 'phase_name']
        unique_together = [('institute_semester', 'phase_name')]
        verbose_name = 'Institute mark at-risk threshold (phase)'
        verbose_name_plural = 'Institute mark at-risk thresholds (phase)'

    def __str__(self):
        return f'{self.institute_semester} {self.phase_name} < {self.fail_below_marks}'


class DepartmentExamPhaseRiskThreshold(models.Model):
    """Department override for mark at-risk cutoff per exam phase (HOD / super admin)."""

    department = models.ForeignKey(
        Department, on_delete=models.CASCADE, related_name='exam_phase_risk_thresholds'
    )
    phase_name = models.CharField(max_length=50)
    fail_below_marks = models.DecimalField(
        max_digits=6,
        decimal_places=2,
        default=9,
        help_text='At-risk if marks strictly less than this.',
    )

    class Meta:
        ordering = ['department', 'phase_name']
        unique_together = [('department', 'phase_name')]
        verbose_name = 'Department mark at-risk threshold (phase)'
        verbose_name_plural = 'Department mark at-risk thresholds (phase)'

    def __str__(self):
        return f'{self.department} {self.phase_name} < {self.fail_below_marks}'


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
    portal_canonical = models.ForeignKey(
        'self',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='portal_duplicate_rows',
        help_text='If set, this row is a duplicate person; timetable/attendance for this row is shown when the linked faculty (with login) uses the portal. Set on the non-login row only.',
    )

    class Meta:
        ordering = ['full_name']
        verbose_name_plural = 'Faculties'

    def __str__(self):
        return f"{self.full_name} ({self.short_name})"


class FacultyDepartmentMembership(models.Model):
    """Links one faculty account to additional departments (same person, e.g. SY_3 + SY_4). Primary row remains Faculty.department."""

    faculty = models.ForeignKey(
        Faculty,
        on_delete=models.CASCADE,
        related_name='department_memberships',
    )
    department = models.ForeignKey(
        Department,
        on_delete=models.CASCADE,
        related_name='faculty_memberships',
    )

    class Meta:
        constraints = [
            models.UniqueConstraint(fields=['faculty', 'department'], name='uniq_faculty_department_membership'),
        ]
        verbose_name = 'Faculty department membership'
        verbose_name_plural = 'Faculty department memberships'

    def __str__(self):
        return f'{self.faculty.short_name} → {self.department.name}'


class Student(models.Model):
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    batch = models.ForeignKey(Batch, on_delete=models.CASCADE)
    roll_no = models.CharField(max_length=30)
    branch = models.CharField(
        max_length=80,
        blank=True,
        help_text='Engineering branch code (e.g. CE, IT); can be filled from marksheet upload.',
    )
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


class RiskStudentMentorLog(models.Model):
    """Faculty follow-up for at-risk mentees; shown on admin Risk students Excel export."""

    KIND_ATTENDANCE_WEEK = 'attendance_week'
    KIND_MARKS_SUBJECT = 'marks_subject'
    KIND_INTRODUCTION_CALL = 'introduction_call'
    PHASE_INTRODUCTION_CALL = 'INTRO'
    KIND_CHOICES = (
        (KIND_ATTENDANCE_WEEK, 'Attendance week'),
        (KIND_MARKS_SUBJECT, 'Marks subject'),
        (KIND_INTRODUCTION_CALL, 'Mentorship introduction call'),
    )
    CONTACT_FATHER = 'Father'
    CONTACT_MOTHER = 'Mother'
    CONTACT_OTHER = 'Other'
    CONTACT_CHOICES = (
        (CONTACT_FATHER, 'Father'),
        (CONTACT_MOTHER, 'Mother'),
        (CONTACT_OTHER, 'Other'),
    )

    student = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='risk_mentor_logs')
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='risk_mentor_logs')
    faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE, related_name='risk_mentor_logs_saved')
    kind = models.CharField(max_length=20, choices=KIND_CHOICES)
    phase = models.CharField(
        max_length=8,
        help_text='T1–T4 for attendance/marks rows; use INTRO for mentorship introduction call.',
    )
    week_index = models.PositiveSmallIntegerField(
        null=True,
        blank=True,
        help_text='0-based week within phase; attendance rows only.',
    )
    subject_name = models.CharField(max_length=200, blank=True, help_text='Marks rows only: failed subject name.')
    contact_person = models.CharField(max_length=20, choices=CONTACT_CHOICES, default=CONTACT_FATHER)
    call_date = models.DateField(null=True, blank=True)
    call_time = models.TimeField(null=True, blank=True)
    duration_minutes = models.PositiveSmallIntegerField(
        null=True,
        blank=True,
        help_text='Total call duration in minutes (introduction call rows).',
    )
    remarks = models.TextField(blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-updated_at']
        constraints = [
            models.UniqueConstraint(
                fields=['student', 'kind', 'phase', 'week_index'],
                condition=models.Q(kind='attendance_week'),
                name='uniq_risk_mentor_log_attendance',
            ),
            models.UniqueConstraint(
                fields=['student', 'kind', 'phase', 'subject_name'],
                condition=models.Q(kind='marks_subject'),
                name='uniq_risk_mentor_log_marks',
            ),
            models.UniqueConstraint(
                fields=['student', 'kind', 'phase'],
                condition=models.Q(kind='introduction_call'),
                name='uniq_risk_mentor_log_intro',
            ),
        ]
        verbose_name = 'Risk student mentor log'
        verbose_name_plural = 'Risk student mentor logs'

    def __str__(self):
        return f'{self.student.roll_no} {self.kind} {self.phase}'


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
    location = models.CharField(
        max_length=255,
        blank=True,
        help_text='Where the doubt session takes place (room, lab, online link, etc.).',
    )
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


class DepartmentExamProfile(models.Model):
    """Links exam coordinators to an attendance Department. Parent = SY/FY/TY lead; children = SY_1, SY_2, …"""
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='department_exam_profiles')
    department = models.ForeignKey(
        Department,
        on_delete=models.CASCADE,
        related_name='exam_coordinator_profiles',
        null=True,
        blank=True,
        help_text='Set by the coordinator on first login (not by exam section). Required before phases or sub-units.',
    )
    parent = models.ForeignKey(
        'self',
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name='children',
        help_text='Null for parent coordinator; set to the SY (parent) profile for SY_1-style accounts.',
    )
    subunit_code = models.CharField(
        max_length=40,
        blank=True,
        help_text='For child accounts only, e.g. SY_1 (must match values in uploaded supervision sheets).',
    )
    invited_by = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='exam_profiles_invited',
        help_text='If set, this coordinator login was created by a hub account (examsy-style).',
    )
    is_hub_coordinator = models.BooleanField(
        default=False,
        help_text='If true, no single linked department: manage delegates, sub-units, and institute-wide supervision phases.',
    )
    institute_semester = models.ForeignKey(
        InstituteSemester,
        on_delete=models.CASCADE,
        related_name='exam_profiles',
        help_text='Academic period for this login; hub accounts set explicitly, others usually match linked department.',
    )

    class Meta:
        verbose_name = 'Department exam profile'
        verbose_name_plural = 'Department exam profiles'
        constraints = [
            models.UniqueConstraint(
                fields=['user', 'institute_semester'],
                name='uniq_dept_exam_parent_user_semester',
                condition=models.Q(parent__isnull=True),
            ),
            models.UniqueConstraint(
                fields=['parent', 'subunit_code', 'institute_semester'],
                name='uniq_exam_subunit_per_parent',
                condition=models.Q(parent__isnull=False),
            ),
        ]

    def __str__(self):
        dept_label = self.department.name if self.department_id else '—'
        if self.parent_id:
            return f"{self.user.username} → {dept_label} ({self.subunit_code})"
        return f"{self.user.username} → {dept_label} (parent)"


class SupervisionExamPhase(models.Model):
    """Supervision phase: either per-department or hub-wide (hub_coordinator set, department null)."""
    institute_semester = models.ForeignKey(
        InstituteSemester,
        on_delete=models.CASCADE,
        related_name='supervision_phases',
    )
    department = models.ForeignKey(
        Department,
        on_delete=models.CASCADE,
        related_name='supervision_phases',
        null=True,
        blank=True,
    )
    hub_coordinator = models.ForeignKey(
        User,
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name='hub_supervision_phases',
        help_text='If set, phase is institute-wide; faculty rows match across all departments.',
    )
    name = models.CharField(max_length=50)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='supervision_phases_created')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']
        constraints = [
            models.UniqueConstraint(
                fields=['department', 'name'],
                condition=models.Q(department__isnull=False),
                name='uniq_supervision_phase_per_dept_name',
            ),
            models.UniqueConstraint(
                fields=['hub_coordinator', 'name', 'institute_semester'],
                condition=models.Q(hub_coordinator__isnull=False, department__isnull=True),
                name='uniq_supervision_phase_per_hub_name',
            ),
        ]
        verbose_name = 'Supervision exam phase'
        verbose_name_plural = 'Supervision exam phases'

    def __str__(self):
        if self.department_id:
            return f'{self.name} ({self.department.name})'
        if self.hub_coordinator_id:
            return f'{self.name} (hub: {self.hub_coordinator.username})'
        return self.name


class SupervisionDuty(models.Model):
    """One supervision slot for a faculty member (from uploaded combined sheet)."""
    OPEN = 'open'
    COMPLETED = 'completed'
    COMPLETION_CHOICES = [
        (OPEN, 'Open'),
        (COMPLETED, 'Completed'),
    ]

    phase = models.ForeignKey(SupervisionExamPhase, on_delete=models.CASCADE, related_name='duties')
    faculty = models.ForeignKey(
        Faculty,
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name='supervision_duties',
    )
    original_faculty = models.ForeignKey(
        Faculty,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='supervision_duties_as_original',
        help_text='Sheet assignee; unchanged when duty is marked proxy to someone else.',
    )
    is_proxy = models.BooleanField(
        default=False,
        help_text='True if a sub-unit coordinator reassigned this duty to another faculty.',
    )
    completion_status = models.CharField(
        max_length=20,
        choices=COMPLETION_CHOICES,
        default=OPEN,
    )
    block_no = models.CharField(max_length=40, blank=True)
    room_no = models.CharField(max_length=40, blank=True)
    completed_at = models.DateTimeField(null=True, blank=True)
    faculty_name_raw = models.CharField(max_length=250, blank=True)
    faculty_short_raw = models.CharField(max_length=40, blank=True)
    supervision_date = models.DateField()
    time_slot = models.CharField(max_length=120, blank=True)
    subject_name = models.CharField(max_length=200, blank=True)
    division_code = models.CharField(
        max_length=40,
        blank=True,
        help_text='Stream/classroom assignment from sheet (e.g. SY_1).',
    )

    class Meta:
        ordering = ['supervision_date', 'time_slot', 'faculty__full_name']
        verbose_name = 'Supervision duty'
        verbose_name_plural = 'Supervision duties'

    def __str__(self):
        who = self.faculty.full_name if self.faculty_id else (self.faculty_name_raw or '?')
        return f"{who} {self.supervision_date} {self.phase.name}"


class PaperCheckingPhase(models.Model):
    """Evaluation / paper checking round: institute (exam section), hub, or single department."""
    institute_semester = models.ForeignKey(
        InstituteSemester,
        on_delete=models.CASCADE,
        related_name='paper_checking_phases',
        null=True,
        blank=True,
        help_text='Empty for exam-section (institute) phases — shared across all academic semesters. '
        'Required for department / hub phases.',
    )
    institute_scope = models.BooleanField(
        default=False,
        help_text='If true, managed by exam section; faculty matched across all departments.',
    )
    department = models.ForeignKey(
        Department,
        on_delete=models.CASCADE,
        related_name='paper_checking_phases',
        null=True,
        blank=True,
    )
    hub_coordinator = models.ForeignKey(
        User,
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name='hub_paper_checking_phases',
    )
    name = models.CharField(max_length=80)
    created_by = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True, related_name='paper_checking_phases_created'
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']
        constraints = [
            models.UniqueConstraint(
                fields=['department', 'name'],
                condition=models.Q(department__isnull=False, institute_scope=False),
                name='uniq_paper_check_phase_dept_name',
            ),
            models.UniqueConstraint(
                fields=['hub_coordinator', 'name', 'institute_semester'],
                condition=models.Q(hub_coordinator__isnull=False, institute_scope=False),
                name='uniq_paper_check_phase_hub_name',
            ),
            models.UniqueConstraint(
                fields=['name'],
                condition=models.Q(institute_scope=True),
                name='uniq_paper_check_phase_institute_name',
            ),
            models.CheckConstraint(
                check=models.Q(institute_scope=True)
                | models.Q(institute_semester__isnull=False),
                name='paper_check_phase_semester_if_dept_or_hub',
            ),
        ]
        verbose_name = 'Paper checking phase'
        verbose_name_plural = 'Paper checking phases'

    def __str__(self):
        if self.institute_scope:
            return f'{self.name} (institute)'
        if self.department_id:
            return f'{self.name} ({self.department.name})'
        if self.hub_coordinator_id:
            return f'{self.name} (hub)'
        return self.name


class PaperCheckingDuty(models.Model):
    """One evaluator row from the evaluation-duty Excel (subject + date + blocks + paper count)."""
    PRACTICAL_ONLINE = 'online'
    PRACTICAL_OFFLINE = 'offline'
    PRACTICAL_MODE_CHOICES = [
        ('', '—'),
        (PRACTICAL_ONLINE, 'Online'),
        (PRACTICAL_OFFLINE, 'Offline'),
    ]
    phase = models.ForeignKey(PaperCheckingPhase, on_delete=models.CASCADE, related_name='duties')
    faculty = models.ForeignKey(
        Faculty,
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name='paper_checking_duties',
    )
    faculty_name_raw = models.CharField(max_length=250, blank=True)
    evaluator_short_raw = models.CharField(max_length=40, blank=True)
    exam_date = models.DateField()
    subject_name = models.CharField(max_length=200, blank=True)
    total_students = models.PositiveIntegerField(default=0)
    deadline_date = models.DateField(
        help_text='Marks / evaluation target date (default: day after exam date).',
    )
    practical_evaluation_mode = models.CharField(
        max_length=16,
        blank=True,
        choices=PRACTICAL_MODE_CHOICES,
        help_text='Legacy field; practical credits now use online+offline rates × papers in one submission.',
    )

    class Meta:
        ordering = ['exam_date', 'subject_name', 'evaluator_short_raw']
        verbose_name = 'Paper checking duty'
        verbose_name_plural = 'Paper checking duties'

    def __str__(self):
        who = self.faculty.short_name if self.faculty_id else (self.evaluator_short_raw or '?')
        return f'{who} {self.exam_date} {self.subject_name}'


class PaperCheckingSubjectCredit(models.Model):
    """Per-phase subject: marks theory vs practical; credits come from piecewise institute rules in code."""
    phase = models.ForeignKey(
        PaperCheckingPhase,
        on_delete=models.CASCADE,
        related_name='subject_credits',
    )
    subject_name = models.CharField(max_length=200)
    is_practical = models.BooleanField(
        default=False,
        help_text='If true, use online/offline credit rates instead of theory credit.',
    )
    credit_per_paper_theory = models.DecimalField(
        max_digits=10,
        decimal_places=4,
        default=0,
        help_text='Legacy field; computation uses institute piecewise formulas (not this value).',
    )
    credit_online_per_paper = models.DecimalField(
        max_digits=10,
        decimal_places=4,
        null=True,
        blank=True,
    )
    credit_offline_per_paper = models.DecimalField(
        max_digits=10,
        decimal_places=4,
        null=True,
        blank=True,
    )
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['subject_name']
        constraints = [
            models.UniqueConstraint(
                fields=['phase', 'subject_name'],
                name='uniq_paper_check_subject_credit_phase_subject',
            ),
        ]
        verbose_name = 'Paper checking subject credit'
        verbose_name_plural = 'Paper checking subject credits'

    def __str__(self):
        return f'{self.phase.name}: {self.subject_name}'


class PaperCheckingAllocation(models.Model):
    """Department + block range from the SY1–SY4 columns for one duty row."""
    duty = models.ForeignKey(PaperCheckingDuty, on_delete=models.CASCADE, related_name='allocations')
    department = models.ForeignKey(
        Department,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='paper_checking_allocations',
    )
    department_code_raw = models.CharField(max_length=40)
    block_range = models.CharField(max_length=160, blank=True)

    class Meta:
        verbose_name = 'Paper checking allocation'
        verbose_name_plural = 'Paper checking allocations'

    def __str__(self):
        return f'{self.department_code_raw} {self.block_range}'


class PaperCheckingAdjustedShare(models.Model):
    """Sub-unit coordinator splits one imported duty row among faculty (sheet row + total unchanged)."""
    duty = models.ForeignKey(
        PaperCheckingDuty,
        on_delete=models.CASCADE,
        related_name='adjusted_shares',
    )
    faculty = models.ForeignKey(
        Faculty,
        on_delete=models.CASCADE,
        related_name='paper_checking_adjusted_shares',
    )
    paper_count = models.PositiveIntegerField()
    created_by = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='paper_check_adjusted_shares_created',
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['duty', 'faculty'],
                name='uniq_paper_check_adjusted_share_duty_faculty',
            ),
        ]
        ordering = ['duty_id', 'faculty__full_name']
        verbose_name = 'Paper checking adjusted share'
        verbose_name_plural = 'Paper checking adjusted shares'

    def __str__(self):
        return f'{self.duty_id} → {self.faculty.short_name}: {self.paper_count}'


class PaperCheckingCompletionRequest(models.Model):
    """Faculty marks paper checking done; sub-unit coordinator approves or dismisses."""
    PENDING = 'pending'
    APPROVED = 'approved'
    REJECTED = 'rejected'
    STATUS_CHOICES = [
        (PENDING, 'Pending'),
        (APPROVED, 'Approved'),
        (REJECTED, 'Dismissed'),
    ]

    duty = models.ForeignKey(
        PaperCheckingDuty,
        on_delete=models.CASCADE,
        related_name='completion_requests',
    )
    faculty = models.ForeignKey(
        Faculty,
        on_delete=models.CASCADE,
        related_name='paper_checking_completion_requests',
    )
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default=PENDING)
    submitted_at = models.DateTimeField(auto_now_add=True)
    decided_at = models.DateTimeField(null=True, blank=True)
    decided_by = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='paper_checking_completion_decisions',
    )

    class Meta:
        ordering = ['-submitted_at']
        constraints = [
            models.UniqueConstraint(
                fields=['duty', 'faculty'],
                condition=models.Q(status='pending'),
                name='uniq_pending_paper_check_completion_per_duty_faculty',
            ),
        ]
        verbose_name = 'Paper checking completion request'
        verbose_name_plural = 'Paper checking completion requests'

    def __str__(self):
        return f'{self.faculty_id} → duty {self.duty_id} ({self.status})'

    def papers_for_faculty_display(self) -> int:
        """Paper count this faculty is responsible for (adjusted share or full duty)."""
        cache = getattr(self, '_papers_for_faculty_display_cache', None)
        if cache is not None:
            return cache
        duty = self.duty
        if (
            getattr(duty, '_prefetched_objects_cache', None)
            and 'adjusted_shares' in duty._prefetched_objects_cache
        ):
            for sh in duty.adjusted_shares.all():
                if sh.faculty_id == self.faculty_id:
                    self._papers_for_faculty_display_cache = sh.paper_count
                    return sh.paper_count
            self._papers_for_faculty_display_cache = duty.total_students
            return self._papers_for_faculty_display_cache
        sh = PaperCheckingAdjustedShare.objects.filter(
            duty_id=self.duty_id, faculty_id=self.faculty_id
        ).first()
        v = sh.paper_count if sh else duty.total_students
        self._papers_for_faculty_display_cache = v
        return v


class PaperSettingPhase(models.Model):
    """Paper setting duty round (separate uploads from checking)."""
    institute_semester = models.ForeignKey(
        InstituteSemester,
        on_delete=models.CASCADE,
        related_name='paper_setting_phases',
        null=True,
        blank=True,
        help_text='Empty for exam-section (institute) phases — shared across all academic semesters. '
        'Required for department / hub phases.',
    )
    institute_scope = models.BooleanField(default=False)
    department = models.ForeignKey(
        Department,
        on_delete=models.CASCADE,
        related_name='paper_setting_phases',
        null=True,
        blank=True,
    )
    hub_coordinator = models.ForeignKey(
        User,
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name='hub_paper_setting_phases',
    )
    name = models.CharField(max_length=80)
    created_by = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True, related_name='paper_setting_phases_created'
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']
        constraints = [
            models.UniqueConstraint(
                fields=['department', 'name'],
                condition=models.Q(department__isnull=False, institute_scope=False),
                name='uniq_paper_set_phase_dept_name',
            ),
            models.UniqueConstraint(
                fields=['hub_coordinator', 'name', 'institute_semester'],
                condition=models.Q(hub_coordinator__isnull=False, institute_scope=False),
                name='uniq_paper_set_phase_hub_name',
            ),
            models.UniqueConstraint(
                fields=['name'],
                condition=models.Q(institute_scope=True),
                name='uniq_paper_set_phase_institute_name',
            ),
            models.CheckConstraint(
                check=models.Q(institute_scope=True)
                | models.Q(institute_semester__isnull=False),
                name='paper_set_phase_semester_if_dept_or_hub',
            ),
        ]
        verbose_name = 'Paper setting phase'
        verbose_name_plural = 'Paper setting phases'

    def __str__(self):
        if self.institute_scope:
            return f'{self.name} (institute)'
        if self.department_id:
            return f'{self.name} ({self.department.name})'
        if self.hub_coordinator_id:
            return f'{self.name} (hub)'
        return self.name


class PaperSettingDuty(models.Model):
    phase = models.ForeignKey(PaperSettingPhase, on_delete=models.CASCADE, related_name='duties')
    faculty = models.ForeignKey(
        Faculty,
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name='paper_setting_duties',
    )
    faculty_name_raw = models.CharField(max_length=250, blank=True)
    faculty_short_raw = models.CharField(max_length=40, blank=True)
    duty_date = models.DateField(null=True, blank=True)
    deadline_date = models.DateField(
        null=True,
        blank=True,
        help_text='Target completion date (from upload or duty date).',
    )
    subject_name = models.CharField(max_length=200, blank=True)
    notes = models.TextField(blank=True)

    class Meta:
        ordering = ['duty_date', 'subject_name', 'faculty_short_raw']
        verbose_name = 'Paper setting duty'
        verbose_name_plural = 'Paper setting duties'

    def __str__(self):
        who = self.faculty.short_name if self.faculty_id else (self.faculty_short_raw or '?')
        return f'{who} {self.subject_name}'


class PaperSettingCompletionRequest(models.Model):
    """Faculty marks paper setting done; department / sub-unit coordinator approves."""

    PENDING = 'pending'
    APPROVED = 'approved'
    REJECTED = 'rejected'
    STATUS_CHOICES = [
        (PENDING, 'Pending'),
        (APPROVED, 'Approved'),
        (REJECTED, 'Dismissed'),
    ]

    duty = models.ForeignKey(
        PaperSettingDuty,
        on_delete=models.CASCADE,
        related_name='completion_requests',
    )
    faculty = models.ForeignKey(
        Faculty,
        on_delete=models.CASCADE,
        related_name='paper_setting_completion_requests',
    )
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default=PENDING)
    submitted_at = models.DateTimeField(auto_now_add=True)
    decided_at = models.DateTimeField(null=True, blank=True)
    decided_by = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='paper_setting_completion_decisions',
    )

    class Meta:
        ordering = ['-submitted_at']
        constraints = [
            models.UniqueConstraint(
                fields=['duty', 'faculty'],
                condition=models.Q(status='pending'),
                name='uniq_pending_paper_set_completion_per_duty_faculty',
            ),
        ]
        verbose_name = 'Paper setting completion request'
        verbose_name_plural = 'Paper setting completion requests'

    def __str__(self):
        return f'{self.faculty_id} → paper setting duty {self.duty_id} ({self.status})'


class DepartmentExamCreditRule(models.Model):
    """Phase-wise credits. department NULL = institute default (all departments). Blank subject = bucket default."""

    TASK_PAPER_SETTING = 'paper_setting'
    TASK_SUPERVISION = 'supervision'
    TASK_PAPER_CHECKING = 'paper_checking'
    TASK_CHOICES = [
        (TASK_PAPER_SETTING, 'Paper setting (per approved duty)'),
        (TASK_SUPERVISION, 'Supervision (per completed duty)'),
        (TASK_PAPER_CHECKING, 'Paper checking (theory fallback: credit per paper if no subject rule)'),
    ]

    BUCKET_T1_T3 = 't1_t3'
    BUCKET_SEE = 'see'
    BUCKET_REMEDIAL = 'remedial'
    BUCKET_FAST_TRACK = 'fast_track'
    BUCKET_CHOICES = [
        (BUCKET_T1_T3, 'T1–T3 / internal'),
        (BUCKET_SEE, 'T4 / SEE'),
        (BUCKET_REMEDIAL, 'Remedial'),
        (BUCKET_FAST_TRACK, 'Fast track'),
    ]

    department = models.ForeignKey(
        Department,
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name='exam_credit_rules',
        help_text='Empty = institute default applied to all departments (unless a department has its own rule).',
    )
    task = models.CharField(max_length=24, choices=TASK_CHOICES)
    phase_bucket = models.CharField(max_length=20, choices=BUCKET_CHOICES)
    subject_name = models.CharField(
        max_length=200,
        blank=True,
        help_text='Leave empty for department default; otherwise overrides for that subject only.',
    )
    credit = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    remuneration = models.DecimalField(
        max_digits=12,
        decimal_places=2,
        default=0,
        help_text='₹ per same unit as credit (per duty for setting/supervision; per paper for checking fallback).',
    )
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['department_id', 'task', 'phase_bucket', 'subject_name']
        constraints = [
            models.UniqueConstraint(
                fields=['department', 'task', 'phase_bucket', 'subject_name'],
                condition=models.Q(department__isnull=False),
                name='uniq_dept_exam_credit_rule',
            ),
            models.UniqueConstraint(
                fields=['task', 'phase_bucket', 'subject_name'],
                condition=models.Q(department__isnull=True),
                name='uniq_institute_exam_credit_rule',
            ),
        ]
        verbose_name = 'Department exam credit rule'
        verbose_name_plural = 'Department exam credit rules'

    def __str__(self):
        subj = (self.subject_name or '').strip() or '(default)'
        who = self.department.name if self.department_id else 'All departments'
        return f'{who} {self.task} {self.phase_bucket} {subj} → {self.credit}'
