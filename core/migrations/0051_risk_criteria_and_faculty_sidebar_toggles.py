import django.db.models.deletion
from decimal import Decimal
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0050_mentor_student_introduction_call'),
    ]

    operations = [
        migrations.AddField(
            model_name='institutesemester',
            name='risk_attendance_min_percent',
            field=models.DecimalField(
                decimal_places=2,
                default=Decimal('75'),
                help_text='Institute default: cumulative attendance below this % is at-risk (all term phases). Departments can override.',
                max_digits=5,
            ),
        ),
        migrations.AddField(
            model_name='department',
            name='faculty_show_mark_attendance',
            field=models.BooleanField(
                default=True,
                help_text='If off, faculty sidebar hides Mark Attendance and saving attendance is blocked.',
            ),
        ),
        migrations.AddField(
            model_name='department',
            name='faculty_show_mentorship',
            field=models.BooleanField(
                default=True,
                help_text='If off, faculty sidebar hides Mentorship Students.',
            ),
        ),
        migrations.AddField(
            model_name='department',
            name='risk_attendance_min_percent',
            field=models.DecimalField(
                blank=True,
                decimal_places=2,
                help_text='If set, overrides institute semester default for at-risk attendance %.',
                max_digits=5,
                null=True,
            ),
        ),
        migrations.CreateModel(
            name='InstituteExamPhaseRiskThreshold',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('phase_name', models.CharField(help_text='Exam phase name, e.g. T1, T2, SEE', max_length=50)),
                (
                    'fail_below_marks',
                    models.DecimalField(
                        decimal_places=2,
                        default=Decimal('9'),
                        help_text='At-risk if marks obtained is strictly less than this (same as legacy < 9).',
                        max_digits=6,
                    ),
                ),
                (
                    'institute_semester',
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name='exam_phase_risk_thresholds',
                        to='core.institutesemester',
                    ),
                ),
            ],
            options={
                'verbose_name': 'Institute mark at-risk threshold (phase)',
                'verbose_name_plural': 'Institute mark at-risk thresholds (phase)',
                'ordering': ['institute_semester', 'phase_name'],
                'unique_together': {('institute_semester', 'phase_name')},
            },
        ),
        migrations.CreateModel(
            name='DepartmentExamPhaseRiskThreshold',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('phase_name', models.CharField(max_length=50)),
                (
                    'fail_below_marks',
                    models.DecimalField(
                        decimal_places=2,
                        default=Decimal('9'),
                        help_text='At-risk if marks strictly less than this.',
                        max_digits=6,
                    ),
                ),
                (
                    'department',
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name='exam_phase_risk_thresholds',
                        to='core.department',
                    ),
                ),
            ],
            options={
                'verbose_name': 'Department mark at-risk threshold (phase)',
                'verbose_name_plural': 'Department mark at-risk thresholds (phase)',
                'ordering': ['department', 'phase_name'],
                'unique_together': {('department', 'phase_name')},
            },
        ),
    ]
