# Exam Admin role + default examadmin user (read-only marks analytics across departments).

from django.db import migrations, models
from django.contrib.auth.hashers import make_password


def create_exam_admin_user(apps, schema_editor):
    User = apps.get_model('auth', 'User')
    UserRole = apps.get_model('accounts', 'UserRole')
    if User.objects.filter(username='examadmin').exists():
        u = User.objects.get(username='examadmin')
    else:
        u = User.objects.create(
            username='examadmin',
            password=make_password('ljiet@123'),
            email='',
            is_active=True,
            is_staff=False,
            is_superuser=False,
        )
    UserRole.objects.update_or_create(
        user=u,
        defaults={'role': 'exam_admin', 'department_id': None},
    )


def remove_exam_admin_user(apps, schema_editor):
    User = apps.get_model('auth', 'User')
    User.objects.filter(username='examadmin').delete()


class Migration(migrations.Migration):

    dependencies = [
        ('accounts', '0003_add_hod_and_week_lock'),
    ]

    operations = [
        migrations.AlterField(
            model_name='userrole',
            name='role',
            field=models.CharField(choices=[
                ('admin', 'Admin'),
                ('hod', 'HOD'),
                ('exam_admin', 'Exam Admin'),
                ('faculty', 'Faculty'),
                ('student', 'Student'),
            ], max_length=20),
        ),
        migrations.RunPython(create_exam_admin_user, remove_exam_admin_user),
    ]
