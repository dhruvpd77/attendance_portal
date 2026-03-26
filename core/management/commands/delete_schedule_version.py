"""
Delete all ScheduleSlot entries for a given department and effective_from date.
Use when an erroneous timetable version was uploaded and you want to remove it.
Example: python manage.py delete_schedule_version --dept=1 --date=2026-03-16
"""
from django.core.management.base import BaseCommand
from datetime import datetime

from core.models import ScheduleSlot, Department


class Command(BaseCommand):
    help = 'Delete schedule slots for a department and effective_from date'

    def add_arguments(self, parser):
        parser.add_argument('--dept', type=int, help='Department ID')
        parser.add_argument('--date', required=True, help='effective_from date (YYYY-MM-DD)')

    def handle(self, *args, **options):
        dept_id = options.get('dept')
        date_str = options['date']
        try:
            eff_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        except ValueError:
            self.stderr.write(self.style.ERROR(f'Invalid date: {date_str}. Use YYYY-MM-DD'))
            return
        dept = None
        if dept_id:
            dept = Department.objects.filter(pk=dept_id).first()
            if not dept:
                self.stderr.write(self.style.ERROR(f'Department {dept_id} not found'))
                return
        qs = ScheduleSlot.objects.filter(effective_from=eff_date)
        if dept:
            qs = qs.filter(department=dept)
        count = qs.count()
        if count == 0:
            self.stdout.write(f'No slots with effective_from={eff_date}' + (f' for dept {dept.name}' if dept else ''))
            return
        qs.delete()
        self.stdout.write(self.style.SUCCESS(f'Deleted {count} slot(s) with effective_from={eff_date}'))
