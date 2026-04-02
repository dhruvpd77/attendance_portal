from django.db import migrations


class Migration(migrations.Migration):
    dependencies = [
        ('core', '0041_institute_semester_exam_faculty_history_cutoff'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='institutesemester',
            name='exam_faculty_portal_history_through_date',
        ),
    ]
