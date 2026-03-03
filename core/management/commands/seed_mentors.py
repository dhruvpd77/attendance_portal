"""
Assign random mentors (from faculty) to students who don't have one.
Mentors are from the same department as the student.

Usage:
  python manage.py seed_mentors
  python manage.py seed_mentors --department "IT"  # optional: department name
"""
import random
from django.core.management.base import BaseCommand
from core.models import Department, Student, Faculty


class Command(BaseCommand):
    help = 'Assign random mentors (from faculty) to students without a mentor'

    def add_arguments(self, parser):
        parser.add_argument(
            '--department', type=str, default=None,
            help='Department name (uses all departments if not set)',
        )

    def handle(self, *args, **options):
        dept_name = options.get('department')
        if dept_name:
            depts = Department.objects.filter(name__iexact=dept_name)
        else:
            depts = Department.objects.all()
        total = 0
        for dept in depts:
            faculties = list(Faculty.objects.filter(department=dept))
            if not faculties:
                self.stdout.write(self.style.WARNING(f'No faculty in {dept.name}. Skipping.'))
                continue
            students = Student.objects.filter(department=dept)
            for s in students:
                s.mentor = random.choice(faculties)
                s.save()
                total += 1
        self.stdout.write(self.style.SUCCESS(f'Assigned mentors to {total} students.'))
