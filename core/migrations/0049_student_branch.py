from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0048_student_user_multi_enrollment'),
    ]

    operations = [
        migrations.AddField(
            model_name='student',
            name='branch',
            field=models.CharField(
                blank=True,
                help_text='Engineering branch code (e.g. CE, IT); can be filled from marksheet upload.',
                max_length=80,
            ),
        ),
    ]
