"""
Send attendance notifications:
- Students: Email when attendance drops below 75% (max once per week)
- Faculty: Reminder before class to mark attendance (daily, to faculty with classes today)
- Admins: Weekly summary of attendance trends

Usage:
  python manage.py send_attendance_notifications
  python manage.py send_attendance_notifications --student-only
  python manage.py send_attendance_notifications --faculty-only
  python manage.py send_attendance_notifications --admin-only
"""
from collections import defaultdict
from datetime import datetime, timedelta

from django.core.management.base import BaseCommand
from django.core.mail import send_mail
from django.conf import settings
from django.utils import timezone

from core.models import (
    Department, Batch, Student, Faculty, ScheduleSlot,
    TermPhase, FacultyAttendance, PhaseHoliday, AttendanceNotificationLog,
)
from accounts.models import UserRole


def get_phase_holidays(dept, phase):
    if not dept or not phase:
        return set()
    return set(
        PhaseHoliday.objects.filter(department=dept, phase=phase.upper()).values_list('date', flat=True)
    )


def get_faculty_subject_for_slot(date, batch, time_slot):
    from core.views import get_faculty_subject_for_slot as _get
    return _get(date, batch, time_slot)


def _compile_phase_weeks(dept, phase):
    tp = TermPhase.objects.filter(department=dept).first()
    if not tp:
        return []
    start = getattr(tp, f'{phase.lower()}_start', None)
    end = getattr(tp, f'{phase.lower()}_end', None)
    if not start or not end:
        return []
    days_set = set(ScheduleSlot.objects.filter(department=dept).values_list('day', flat=True).distinct())
    days_set = {d.lower() for d in days_set if d}
    holidays = get_phase_holidays(dept, phase)
    dates = []
    cur = start
    while cur <= end:
        if cur not in holidays and cur.strftime('%A').lower() in days_set:
            dates.append(cur)
        cur += timedelta(days=1)
    dates = sorted(dates)
    weeks = []
    week = []
    last_w = None
    for d in dates:
        w = d.isocalendar()[1]
        if last_w is not None and w != last_w and week:
            weeks.append(week)
            week = []
        week.append(d)
        last_w = w
    if week:
        weeks.append(week)
    return weeks


def student_low_attendance_emails(command):
    """Send email to students with attendance below 75%. Max once per week."""
    sent = 0
    for dept in Department.objects.all():
        tp = TermPhase.objects.filter(department=dept).first()
        phase = None
        for p in ['T1', 'T2', 'T3', 'T4']:
            if tp and getattr(tp, f'{p.lower()}_start', None) and getattr(tp, f'{p.lower()}_end', None):
                phase = p
                break
        if not phase:
            continue
        weeks = _compile_phase_weeks(dept, phase)
        all_dates = set()
        for w in weeks:
            all_dates.update(w)
        all_dates = sorted(all_dates)
        if not all_dates:
            continue
        batches = list(Batch.objects.filter(department=dept))
        batch_scheduled = defaultdict(set)
        for batch in batches:
            for d in all_dates:
                weekday = d.strftime('%A')
                for slot in ScheduleSlot.objects.filter(batch=batch, day=weekday).values_list('time_slot', flat=True).distinct():
                    batch_scheduled[batch.id].add((d, slot))
        batch_att_map = defaultdict(lambda: defaultdict(set))
        for batch in batches:
            for att in FacultyAttendance.objects.filter(batch=batch, date__in=all_dates):
                key = (att.date, att.lecture_slot)
                batch_att_map[batch.id][key] = set(x.strip() for x in (att.absent_roll_numbers or '').split(',') if x.strip())
        students = Student.objects.filter(department=dept).select_related('batch')
        for s in students:
            str_roll = str(s.roll_no)
            scheduled = batch_scheduled.get(s.batch_id, set())
            held = len(scheduled)
            attended = sum(1 for (d, slot) in scheduled if (d, slot) in batch_att_map[s.batch_id] and str_roll not in batch_att_map[s.batch_id][(d, slot)])
            pct = round(attended / held * 100, 1) if held else 0
            if held and pct < 75:
                email = s.email or (s.user.email if s.user else None)
                if not email:
                    continue
                last_sent = AttendanceNotificationLog.objects.filter(
                    student=s, notification_type='low_attendance'
                ).order_by('-sent_at').first()
                if last_sent and (timezone.now() - last_sent.sent_at).days < 7:
                    continue
                try:
                    send_mail(
                        subject='LJIET Attendance Alert: Your attendance is below 75%',
                        message=f'Dear {s.name},\n\nYour current attendance is {pct}% (attended {attended} of {held} lectures).\n\nPlease ensure you attend classes regularly to maintain the required 75% attendance.\n\n— LJIET Attendance Portal',
                        from_email=settings.DEFAULT_FROM_EMAIL,
                        recipient_list=[email],
                        fail_silently=True,
                    )
                    AttendanceNotificationLog.objects.create(student=s, notification_type='low_attendance')
                    sent += 1
                except Exception as e:
                    command.stdout.write(command.style.WARNING(f'Failed to email {s.roll_no}: {e}'))
    return sent


def faculty_reminder_emails(command):
    """Send reminder to faculty who have classes today."""
    today = timezone.localdate()
    weekday = today.strftime('%A')
    faculties_with_slots = Faculty.objects.filter(
        scheduleslot__day=weekday
    ).distinct().select_related('department')
    sent = 0
    for fac in faculties_with_slots:
        email = fac.email or (fac.user.email if fac.user else None)
        if not email:
            continue
        slots = ScheduleSlot.objects.filter(faculty=fac, day=weekday).select_related('batch', 'subject').order_by('time_slot')
        slot_list = ', '.join(f'{s.batch.name} {s.subject.name} ({s.time_slot})' for s in slots[:5])
        if len(slots) > 5:
            slot_list += f' and {len(slots) - 5} more'
        try:
            send_mail(
                subject=f'Reminder: Mark attendance for today\'s classes — {today.strftime("%d %b %Y")}',
                message=f'Dear {fac.full_name},\n\nYou have classes scheduled for today ({weekday}):\n{slot_list}\n\nPlease mark attendance at: LJIET Attendance Portal\n\n— LJIET Attendance',
                from_email=settings.DEFAULT_FROM_EMAIL,
                recipient_list=[email],
                fail_silently=True,
            )
            sent += 1
        except Exception as e:
            command.stdout.write(command.style.WARNING(f'Failed to email faculty {fac.short_name}: {e}'))
    return sent


def admin_weekly_summary(command):
    """Send weekly attendance summary to admins."""
    admin_users = UserRole.objects.filter(role='admin').select_related('user', 'department')
    if not admin_users.exists():
        admin_users = UserRole.objects.filter(role='admin')
    dept_ids = set(ur.department_id for ur in admin_users if ur.department_id)
    if not dept_ids:
        dept_ids = set(Department.objects.values_list('id', flat=True))
    lines = ['LJIET Attendance — Weekly Summary', '=' * 40, '']
    for dept in Department.objects.filter(id__in=dept_ids):
        tp = TermPhase.objects.filter(department=dept).first()
        phase = 'T1'
        for p in ['T1', 'T2', 'T3', 'T4']:
            if tp and getattr(tp, f'{p.lower()}_start', None):
                phase = p
                break
        weeks = _compile_phase_weeks(dept, phase)
        if not weeks:
            lines.append(f'{dept.name}: No term phases set.')
            continue
        all_dates = set()
        for w in weeks:
            all_dates.update(w)
        batches = list(Batch.objects.filter(department=dept))
        batch_scheduled = defaultdict(set)
        for batch in batches:
            for d in all_dates:
                weekday = d.strftime('%A')
                for slot in ScheduleSlot.objects.filter(batch=batch, day=weekday).values_list('time_slot', flat=True).distinct():
                    batch_scheduled[batch.id].add((d, slot))
        batch_att_map = defaultdict(lambda: defaultdict(set))
        for batch in batches:
            for att in FacultyAttendance.objects.filter(batch=batch, date__in=all_dates):
                key = (att.date, att.lecture_slot)
                batch_att_map[batch.id][key] = set(x.strip() for x in (att.absent_roll_numbers or '').split(',') if x.strip())
        students = Student.objects.filter(department=dept).select_related('batch')
        total_held = total_attended = at_risk = 0
        for s in students:
            str_roll = str(s.roll_no)
            scheduled = batch_scheduled.get(s.batch_id, set())
            held = len(scheduled)
            attended = sum(1 for (d, slot) in scheduled if (d, slot) in batch_att_map[s.batch_id] and str_roll not in batch_att_map[s.batch_id][(d, slot)])
            total_held += held
            total_attended += attended
            if held and round(attended / held * 100, 1) < 75:
                at_risk += 1
        avg_pct = round(total_attended / total_held * 100, 1) if total_held else 0
        lines.append(f'{dept.name}: {total_attended}/{total_held} overall ({avg_pct}%), {at_risk} at-risk students.')
    body = '\n'.join(lines)
    recipients = list(set(ur.user.email for ur in admin_users if ur.user and ur.user.email))
    if not recipients:
        return 0
    try:
        send_mail(
            subject='LJIET Attendance — Weekly Summary',
            message=body,
            from_email=settings.DEFAULT_FROM_EMAIL,
            recipient_list=recipients,
            fail_silently=True,
        )
        return len(recipients)
    except Exception as e:
        command.stdout.write(command.style.WARNING(f'Failed to send admin summary: {e}'))
        return 0


class Command(BaseCommand):
    help = 'Send attendance notifications (student low-attendance, faculty reminder, admin weekly summary)'

    def add_arguments(self, parser):
        parser.add_argument('--student-only', action='store_true')
        parser.add_argument('--faculty-only', action='store_true')
        parser.add_argument('--admin-only', action='store_true')

    def handle(self, *args, **options):
        student_only = options.get('student_only')
        faculty_only = options.get('faculty_only')
        admin_only = options.get('admin_only')
        if not (student_only or faculty_only or admin_only):
            student_only = faculty_only = admin_only = True
        total = 0
        if student_only:
            n = student_low_attendance_emails(self)
            total += n
            self.stdout.write(self.style.SUCCESS(f'Sent {n} student low-attendance emails.'))
        if faculty_only:
            n = faculty_reminder_emails(self)
            total += n
            self.stdout.write(self.style.SUCCESS(f'Sent {n} faculty reminder emails.'))
        if admin_only:
            n = admin_weekly_summary(self)
            total += n
            self.stdout.write(self.style.SUCCESS(f'Sent admin weekly summary to {n} recipients.'))
        self.stdout.write(self.style.SUCCESS(f'Done. Total notifications: {total}'))
