"""Create or update the exam section login (username examsection by default)."""
from django.core.management.base import BaseCommand
from django.contrib.auth.models import User
from accounts.models import UserRole


class Command(BaseCommand):
    help = 'Ensures exam section user exists with role exam_section (default username: examsection).'

    def add_arguments(self, parser):
        parser.add_argument('--username', default='examsection')
        parser.add_argument('--password', required=True)

    def handle(self, *args, **options):
        username = options['username'].strip()
        password = options['password']
        user, created = User.objects.get_or_create(username=username, defaults={'email': ''})
        if not created and not user.has_usable_password():
            user.set_password(password)
            user.save()
        elif created:
            user.set_password(password)
            user.save()
        else:
            user.set_password(password)
            user.save()
        rp, _ = UserRole.objects.update_or_create(
            user=user,
            defaults={'role': 'exam_section', 'department': None},
        )
        self.stdout.write(self.style.SUCCESS(f'Exam section user "{username}" ready (role={rp.role}).'))
