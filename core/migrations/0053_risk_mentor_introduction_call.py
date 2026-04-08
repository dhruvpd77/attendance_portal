# Generated manually

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0052_facultydoubtrequest_location'),
    ]

    operations = [
        migrations.AlterField(
            model_name='riskstudentmentorlog',
            name='phase',
            field=models.CharField(
                help_text='T1–T4 for attendance/marks rows; use INTRO for mentorship introduction call.',
                max_length=8,
            ),
        ),
        migrations.AddField(
            model_name='riskstudentmentorlog',
            name='duration_minutes',
            field=models.PositiveSmallIntegerField(
                blank=True,
                help_text='Total call duration in minutes (introduction call rows).',
                null=True,
            ),
        ),
        migrations.AddConstraint(
            model_name='riskstudentmentorlog',
            constraint=models.UniqueConstraint(
                condition=models.Q(kind='introduction_call'),
                fields=('student', 'kind', 'phase'),
                name='uniq_risk_mentor_log_intro',
            ),
        ),
    ]
