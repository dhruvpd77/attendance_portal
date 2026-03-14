# Generated manually for attendance lock time feature

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0006_add_absent_reasons'),
    ]

    operations = [
        migrations.CreateModel(
            name='AttendanceLockSetting',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('lock_hour', models.PositiveSmallIntegerField(default=17)),
                ('lock_minute', models.PositiveSmallIntegerField(default=0)),
                ('enabled', models.BooleanField(default=False)),
            ],
            options={
                'verbose_name': 'Attendance lock time',
                'verbose_name_plural': 'Attendance lock time',
            },
        ),
    ]
