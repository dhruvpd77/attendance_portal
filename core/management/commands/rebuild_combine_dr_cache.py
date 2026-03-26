"""
Rebuild FacultyCombineDrCache from FacultyAttendance (DR weekly combine rows).
Run after deploy, timetable bulk changes, or if counts look wrong.

  python manage.py rebuild_combine_dr_cache
  python manage.py rebuild_combine_dr_cache --dept=3
"""
from django.core.management.base import BaseCommand
from django.db.models import Exists, OuterRef

from core.models import Department, FacultyAttendance
from core.views import rebuild_faculty_combine_cache_dept


class Command(BaseCommand):
    help = 'Rebuild faculty DR combine cache from saved attendance'

    def add_arguments(self, parser):
        parser.add_argument('--dept', type=int, help='Department ID only (default: all with attendance)')

    def handle(self, *args, **options):
        dept_id = options.get('dept')
        if dept_id:
            depts = list(Department.objects.filter(pk=dept_id))
            if not depts:
                self.stderr.write(self.style.ERROR(f'Department {dept_id} not found'))
                return
        else:
            att = FacultyAttendance.objects.filter(batch__department_id=OuterRef('pk'))
            depts = list(
                Department.objects.filter(Exists(att)).order_by('name')
            )
            if not depts:
                self.stdout.write('No attendance rows — nothing to rebuild.')
                return
        for d in depts:
            rebuild_faculty_combine_cache_dept(d)
            self.stdout.write(self.style.SUCCESS(f'Rebuilt combine DR cache for {d.name} (id={d.id})'))
