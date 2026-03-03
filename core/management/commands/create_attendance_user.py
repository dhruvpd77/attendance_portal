"""
Create a user with role (admin/faculty/student) and optionally link to Faculty or Student.
Usage:
  python manage.py create_attendance_user admin adminuser pass123
  python manage.py create_attendance_user faculty fac1 pass123 --faculty-id=1
  python manage.py create_attendance_user student stu1 pass123 --student-id=1
"""
from django.core.management.base import BaseCommand
from django.contrib.auth.models import User
from accounts.models import UserRole
from core.models import Faculty, Student


class Command(BaseCommand):
    help = 'Create user with role: admin, faculty, or student'

    def add_arguments(self, parser):
        parser.add_argument('role', choices=['admin', 'faculty', 'student'])
        parser.add_argument('username')
        parser.add_argument('password')
        parser.add_argument('--faculty-id', type=int, help='Faculty PK to link (for role=faculty)')
        parser.add_argument('--student-id', type=int, help='Student PK to link (for role=student)')
        parser.add_argument('--email', default='')

    def handle(self, *args, **options):
        role = options['role']
        username = options['username']
        password = options['password']
        email = options.get('email') or ''
        if User.objects.filter(username=username).exists():
            self.stdout.write(self.style.WARNING(f'User {username} already exists.'))
            return
        user = User.objects.create_user(username=username, password=password, email=email)
        UserRole.objects.create(user=user, role=role)
        if role == 'faculty' and options.get('faculty_id'):
            faculty = Faculty.objects.filter(pk=options['faculty_id']).first()
            if faculty:
                faculty.user = user
                faculty.save()
                self.stdout.write(self.style.SUCCESS(f'User {username} linked to faculty {faculty.full_name}'))
        if role == 'student' and options.get('student_id'):
            student = Student.objects.filter(pk=options['student_id']).first()
            if student:
                student.user = user
                student.save()
                self.stdout.write(self.style.SUCCESS(f'User {username} linked to student {student.name}'))
        self.stdout.write(self.style.SUCCESS(f'User {username} created with role {role}.'))
