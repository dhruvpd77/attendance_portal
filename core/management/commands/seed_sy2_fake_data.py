"""
Seed fake data for department sy2:
- 10 students per batch (for all batches)
- Fake T1 phase attendance for all batches

Creates batches, faculty, subjects, schedule slots, and TermPhase if they don't exist.

Usage:
  python manage.py seed_sy2_fake_data
"""
import random
from datetime import date, timedelta

from django.core.management.base import BaseCommand

from core.models import (
    Department, Batch, Student, Faculty, Subject,
    ScheduleSlot, TermPhase, FacultyAttendance, PhaseHoliday,
)


def get_phase_holidays(dept, phase):
    if not dept or not phase:
        return set()
    return set(
        PhaseHoliday.objects.filter(department=dept, phase=phase.upper()).values_list('date', flat=True)
    )


class Command(BaseCommand):
    help = 'Seed sy2 department: 10 students per batch + fake T1 attendance. Creates batches/faculty/subjects/slots if missing.'

    def handle(self, *args, **options):
        dept = Department.objects.filter(name__iexact='sy2').first()
        if not dept:
            self.stdout.write(self.style.ERROR('Department "sy2" not found. Create it first.'))
            return

        self.stdout.write(f'Seeding department: {dept.name}')

        # --- Ensure batches exist
        batches = list(Batch.objects.filter(department=dept).order_by('name'))
        if not batches:
            for name in ['B1', 'B2', 'B3']:
                b, _ = Batch.objects.get_or_create(department=dept, name=name)
                batches.append(b)
            self.stdout.write(self.style.SUCCESS(f'Created batches: {[b.name for b in batches]}'))
        else:
            self.stdout.write(f'Using existing batches: {[b.name for b in batches]}')

        # --- Ensure faculty exist
        faculty_list = list(Faculty.objects.filter(department=dept).order_by('short_name'))
        if not faculty_list:
            defaults = [
                ('Dr. Faculty One', 'F1', 'f1@sy2.test'),
                ('Dr. Faculty Two', 'F2', 'f2@sy2.test'),
                ('Dr. Faculty Three', 'F3', 'f3@sy2.test'),
            ]
            for full_name, short_name, email in defaults:
                f, _ = Faculty.objects.get_or_create(
                    department=dept, short_name=short_name,
                    defaults={'full_name': full_name, 'email': email}
                )
                faculty_list.append(f)
            self.stdout.write(self.style.SUCCESS(f'Created faculty: {[f.short_name for f in faculty_list]}'))
        else:
            self.stdout.write(f'Using existing faculty: {[f.short_name for f in faculty_list]}')

        # --- Ensure subjects exist
        subjects = list(Subject.objects.filter(department=dept).order_by('name'))
        if not subjects:
            for name, code in [('Mathematics', 'MATH101'), ('Physics', 'PHY101'), ('Programming', 'CS101'), ('English', 'ENG101')]:
                s, _ = Subject.objects.get_or_create(
                    department=dept, name=name,
                    defaults={'code': code}
                )
                subjects.append(s)
            self.stdout.write(self.style.SUCCESS(f'Created subjects: {[s.name for s in subjects]}'))
        else:
            self.stdout.write(f'Using existing subjects: {[s.name for s in subjects]}')

        # --- Ensure schedule slots exist (Mon-Fri, 3 slots per day)
        slots_exist = ScheduleSlot.objects.filter(department=dept).exists()
        if not slots_exist:
            days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
            time_slots = ['08:45-09:45', '10:00-11:00', '11:15-12:15']
            for batch in batches:
                for day in days:
                    for i, ts in enumerate(time_slots):
                        fac = faculty_list[i % len(faculty_list)]
                        subj = subjects[i % len(subjects)]
                        ScheduleSlot.objects.get_or_create(
                            department=dept, batch=batch, day=day, time_slot=ts,
                            defaults={'faculty': fac, 'subject': subj}
                        )
            self.stdout.write(self.style.SUCCESS('Created schedule slots (Mon-Fri, 3 slots/day per batch)'))
        else:
            self.stdout.write('Using existing schedule slots')

        # --- Ensure TermPhase T1 exists
        tp, created = TermPhase.objects.get_or_create(
            department=dept,
            defaults={
                't1_start': date(2026, 3, 9),
                't1_end': date(2026, 3, 28),
            }
        )
        if created:
            self.stdout.write(self.style.SUCCESS(f'Created TermPhase T1: {tp.t1_start} to {tp.t1_end}'))
        elif not tp.t1_start or not tp.t1_end:
            tp.t1_start = date(2026, 3, 9)
            tp.t1_end = date(2026, 3, 28)
            tp.save()
            self.stdout.write(self.style.SUCCESS(f'Set TermPhase T1: {tp.t1_start} to {tp.t1_end}'))
        else:
            self.stdout.write(f'Using TermPhase T1: {tp.t1_start} to {tp.t1_end}')

        # --- Build T1 lecture dates (excluding holidays)
        days_set = set(
            ScheduleSlot.objects.filter(department=dept).values_list('day', flat=True).distinct()
        )
        days_set = {d.lower() for d in days_set if d}
        holidays = get_phase_holidays(dept, 'T1')
        t1_dates = []
        cur = tp.t1_start
        while cur <= tp.t1_end:
            if cur not in holidays and cur.strftime('%A').lower() in days_set:
                t1_dates.append(cur)
            cur += timedelta(days=1)
        t1_dates = sorted(t1_dates)
        self.stdout.write(f'T1 lecture days: {len(t1_dates)} ({t1_dates[0] if t1_dates else "none"} to {t1_dates[-1] if t1_dates else "none"})')

        # --- Add 10 students per batch
        students_added = 0
        for batch in batches:
            existing = set(Student.objects.filter(batch=batch).values_list('roll_no', flat=True))
            needed = 10
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
            students_added += added
            if added:
                self.stdout.write(self.style.SUCCESS(f'  Added {added} students to batch {batch.name}'))

        if students_added:
            self.stdout.write(self.style.SUCCESS(f'Total students added: {students_added}'))
        else:
            self.stdout.write('All batches already have 10+ students.')

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
                    n_absent = random.randint(0, min(4, len(students)))
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
            f'Lecture days: {len(t1_dates)}, Batches: {len(batches)}'
        ))
        self.stdout.write(self.style.SUCCESS('Done. sy2 fake data seeded.'))
