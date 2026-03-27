# Attendance lock time is per-department (was a single global row).

from django.db import migrations, models
import django.db.models.deletion


def migrate_lock_to_departments(apps, schema_editor):
    AttendanceLockSetting = apps.get_model('core', 'AttendanceLockSetting')
    Department = apps.get_model('core', 'Department')
    legacy = AttendanceLockSetting.objects.filter(department__isnull=True).first()
    h, m, e = 17, 0, False
    if legacy:
        h, m, e = legacy.lock_hour, legacy.lock_minute, legacy.enabled
    for d in Department.objects.all():
        AttendanceLockSetting.objects.get_or_create(
            department_id=d.id,
            defaults={'lock_hour': h, 'lock_minute': m, 'enabled': e},
        )
    AttendanceLockSetting.objects.filter(department__isnull=True).delete()


def backwards_noop(apps, schema_editor):
    pass


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0021_faculty_combine_cache_effective_load'),
    ]

    operations = [
        migrations.AddField(
            model_name='attendancelocksetting',
            name='department',
            field=models.OneToOneField(
                null=True,
                blank=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name='attendance_lock_setting',
                to='core.department',
            ),
        ),
        migrations.RunPython(migrate_lock_to_departments, backwards_noop),
        migrations.AlterField(
            model_name='attendancelocksetting',
            name='department',
            field=models.OneToOneField(
                on_delete=django.db.models.deletion.CASCADE,
                related_name='attendance_lock_setting',
                to='core.department',
            ),
        ),
        migrations.AlterModelOptions(
            name='attendancelocksetting',
            options={
                'verbose_name': 'Attendance lock time',
                'verbose_name_plural': 'Attendance lock times',
            },
        ),
    ]
