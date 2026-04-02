"""
Merge one Subject into another in a department: re-point all FKs and string labels, then delete the source subject.

Usage:
  python manage.py merge_subjects --department SY_4 --from "PYTHON-II" --to "FCSP-II"
  python manage.py merge_subjects --department SY_4 --from "PYTHON-II" --to "FCSP-II" --dry-run
"""
from django.core.management.base import BaseCommand, CommandError

from core.models import Department
from core.subject_merge import (
    SubjectMergeError,
    execute_subject_merge,
    resolve_merge_subjects,
    subject_merge_counts,
)


class Command(BaseCommand):
    help = 'Merge source subject into target subject within one department (updates references, deletes source).'

    def add_arguments(self, parser):
        parser.add_argument('--department', required=True, help='Department name or pk (e.g. SY_4)')
        parser.add_argument('--from', dest='from_name', required=True, help='Source subject name to remove')
        parser.add_argument('--to', dest='to_name', required=True, help='Target subject name to keep')
        parser.add_argument('--dry-run', action='store_true', help='Show counts only; no DB writes')

    def handle(self, *args, **options):
        dept_key = (options['department'] or '').strip()
        from_name = (options['from_name'] or '').strip()
        to_name = (options['to_name'] or '').strip()
        dry = options['dry_run']

        dept = Department.objects.filter(name__iexact=dept_key).first()
        if not dept and dept_key.isdigit():
            dept = Department.objects.filter(pk=int(dept_key)).first()
        if not dept:
            raise CommandError(f'Department not found: {dept_key!r}')

        try:
            src, dst, _target_label = resolve_merge_subjects(dept, from_name, to_name)
        except SubjectMergeError as exc:
            raise CommandError(str(exc)) from exc

        counts = subject_merge_counts(dept, src, from_name)
        for key, n in sorted(counts.items()):
            self.stdout.write(f'  {key}: {n}')

        if dry:
            self.stdout.write(self.style.WARNING('Dry run — no changes applied.'))
            return

        stats = execute_subject_merge(dept, from_name, to_name)
        for key, n in sorted(stats.items()):
            self.stdout.write(self.style.SUCCESS(f'{key}: {n}'))
        self.stdout.write(self.style.SUCCESS(f'Done: merged {from_name!r} into {dst.name!r} in {dept.name}.'))
