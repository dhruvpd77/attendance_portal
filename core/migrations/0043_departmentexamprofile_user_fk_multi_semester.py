from django.conf import settings
from django.db import migrations, models
from django.db.models import Q


class Migration(migrations.Migration):
    dependencies = [
        ('core', '0042_remove_exam_faculty_portal_history_date'),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.AlterField(
            model_name='departmentexamprofile',
            name='user',
            field=models.ForeignKey(
                on_delete=models.deletion.CASCADE,
                related_name='department_exam_profiles',
                to=settings.AUTH_USER_MODEL,
            ),
        ),
        migrations.AddConstraint(
            model_name='departmentexamprofile',
            constraint=models.UniqueConstraint(
                condition=Q(parent__isnull=True),
                fields=('user', 'institute_semester'),
                name='uniq_dept_exam_parent_user_semester',
            ),
        ),
    ]
