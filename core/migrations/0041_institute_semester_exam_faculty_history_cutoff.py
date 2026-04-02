from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ('core', '0040_faculty_portal_exam_analytics_toggles'),
    ]

    operations = [
        migrations.AddField(
            model_name='institutesemester',
            name='exam_faculty_portal_history_through_date',
            field=models.DateField(
                blank=True,
                help_text='Optional. When set, faculty “My exam duties” and “Exam credits” treat this date as the end of the archived cycle: rows on or before this date appear only under History; main totals and actions use activity after this date only.',
                null=True,
            ),
        ),
        migrations.AlterModelOptions(
            name='institutesemester',
            options={
                'ordering': ['-sort_order', '-pk'],
                'verbose_name': 'Academic semester',
                'verbose_name_plural': 'Academic semesters',
            },
        ),
    ]
