"""
Seed fake T1 phase attendance and add 5 fake students per batch (skip A3).
Usage (from attendance_portal directory):
  python manage.py seed_t1_attendance
  python manage.py seed_t1_attendance --department "IT"  # optional: department name
"""
import random
from datetime import timedelta
from django.core.management.base import BaseCommand
from core.models import (
    Department, Batch, Student, ScheduleSlot,
    TermPhase, FacultyAttendance,
)


class Command(BaseCommand):
    help = 'Add 5 fake students per batch (except A3) and create T1 phase attendance entries for all batches/subjects'

    def add_arguments(self, parser):
        parser.add_argument(
            '--department', type=str, default=None,
            help='Department name (uses first department if not set)',
        )
        parser.add_argument(
            '--skip-students', action='store_true',
            help='Do not add fake students, only attendance',
        )

    def handle(self, *args, **options):
        dept_name = options.get('department')
        if dept_name:
            dept = Department.objects.filter(name__iexact=dept_name).first()
        else:
            dept = Department.objects.first()
        if not dept:
            self.stdout.write(self.style.ERROR('No department found. Create a department first.'))
            return

        tp = TermPhase.objects.filter(department=dept).first()
        if not tp or not tp.t1_start or not tp.t1_end:
            self.stdout.write(self.style.ERROR('TermPhase T1 start/end not set for this department. Set T1 dates first.'))
            return

        days_set = set(
            ScheduleSlot.objects.filter(department=dept)
            .values_list('day', flat=True).distinct()
        )
        days_set = {d.lower() for d in days_set if d}
        if not days_set:
            self.stdout.write(self.style.ERROR('No schedule slots found. Upload timetable or add schedule first.'))
            return

        # Build T1 lecture dates
        t1_dates = []
        cur = tp.t1_start
        while cur <= tp.t1_end:
            if cur.strftime('%A').lower() in days_set:
                t1_dates.append(cur)
            cur += timedelta(days=1)
        t1_dates = sorted(t1_dates)
        self.stdout.write(f'T1: {tp.t1_start} to {tp.t1_end}, {len(t1_dates)} lecture days')

        batches = list(Batch.objects.filter(department=dept).order_by('name'))
        if not batches:
            self.stdout.write(self.style.ERROR('No batches found.'))
            return

        # --- Add 5 fake students per batch (except A3)
        if not options.get('skip_students'):
            for batch in batches:
                if batch.name.upper() == 'A3':
                    self.stdout.write(f'Skipping students for {batch.name} (already added).')
                    continue
                existing = set(Student.objects.filter(batch=batch).values_list('roll_no', flat=True))
                needed = 5
                added = 0
                for i in range(1, 100):
                    if added >= needed:
                        break
                    roll = f'{batch.name}-F{i}'
                    if roll in existing:
                        continue
                    Student.objects.get_or_create(
                        department=dept,
                        batch=batch,
                        roll_no=roll,
                        defaults={
                            'name': f'Fake Student {i} ({batch.name})',
                            'enrollment_no': f'ENR{batch.id}{i:03d}',
                        },
                    )
                    existing.add(roll)
                    added += 1
                if added:
                    self.stdout.write(self.style.SUCCESS(f'Added {added} fake students to batch {batch.name}'))

        # --- Create T1 attendance for each (date, batch, slot)
        created_count = 0
        updated_count = 0
        for batch in batches:
            students = list(Student.objects.filter(batch=batch).values_list('roll_no', flat=True))
            if not students:
                self.stdout.write(self.style.WARNING(f'No students in batch {batch.name}, skipping attendance.'))
                continue
            for d in t1_dates:
                weekday = d.strftime('%A')
                slots = ScheduleSlot.objects.filter(
                    department=dept, batch=batch, day=weekday
                ).select_related('faculty')
                for slot in slots:
                    # Random 0 to min(3, len(students)) absent per slot for variety
                    n_absent = random.randint(0, min(3, len(students)))
                    absent_rolls = random.sample(students, n_absent) if n_absent else []
                    absent_str = ','.join(str(r) for r in absent_rolls)
                    obj, created = FacultyAttendance.objects.update_or_create(
                        faculty=slot.faculty,
                        date=d,
                        batch=batch,
                        lecture_slot=slot.time_slot,
                        defaults={'absent_roll_numbers': absent_str},
                    )
                    if created:
                        created_count += 1
                    else:
                        updated_count += 1

        self.stdout.write(self.style.SUCCESS(
            f'T1 attendance: {created_count} created, {updated_count} updated. '
            f'Total lecture days: {len(t1_dates)}, Batches: {len(batches)}'
        ))
