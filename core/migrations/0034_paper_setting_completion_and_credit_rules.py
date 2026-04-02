# Generated manually for paper setting workflow and department credit rules.

import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models
from django.db.models import Q


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('core', '0033_paper_phase_institute_optional_semester'),
    ]

    operations = [
        migrations.AddField(
            model_name='papersettingduty',
            name='deadline_date',
            field=models.DateField(
                blank=True,
                help_text='Target completion date (from upload or duty date).',
                null=True,
            ),
        ),
        migrations.CreateModel(
            name='DepartmentExamCreditRule',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('task', models.CharField(
                    choices=[
                        ('paper_setting', 'Paper setting (per approved duty)'),
                        ('supervision', 'Supervision (per completed duty)'),
                    ],
                    max_length=24,
                )),
                ('phase_bucket', models.CharField(
                    choices=[
                        ('t1_t3', 'T1–T3 / internal'),
                        ('see', 'T4 / SEE'),
                        ('remedial', 'Remedial'),
                        ('fast_track', 'Fast track'),
                    ],
                    max_length=20,
                )),
                ('subject_name', models.CharField(
                    blank=True,
                    help_text='Leave empty for department default; otherwise overrides for that subject only.',
                    max_length=200,
                )),
                ('credit', models.DecimalField(decimal_places=2, default=0, max_digits=8)),
                ('updated_at', models.DateTimeField(auto_now=True)),
                ('department', models.ForeignKey(
                    on_delete=django.db.models.deletion.CASCADE,
                    related_name='exam_credit_rules',
                    to='core.department',
                )),
            ],
            options={
                'verbose_name': 'Department exam credit rule',
                'verbose_name_plural': 'Department exam credit rules',
                'ordering': ['department_id', 'task', 'phase_bucket', 'subject_name'],
            },
        ),
        migrations.AddConstraint(
            model_name='departmentexamcreditrule',
            constraint=models.UniqueConstraint(
                fields=('department', 'task', 'phase_bucket', 'subject_name'),
                name='uniq_dept_exam_credit_rule',
            ),
        ),
        migrations.CreateModel(
            name='PaperSettingCompletionRequest',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('status', models.CharField(
                    choices=[('pending', 'Pending'), ('approved', 'Approved'), ('rejected', 'Dismissed')],
                    default='pending',
                    max_length=20,
                )),
                ('submitted_at', models.DateTimeField(auto_now_add=True)),
                ('decided_at', models.DateTimeField(blank=True, null=True)),
                ('decided_by', models.ForeignKey(
                    blank=True,
                    null=True,
                    on_delete=django.db.models.deletion.SET_NULL,
                    related_name='paper_setting_completion_decisions',
                    to=settings.AUTH_USER_MODEL,
                )),
                ('duty', models.ForeignKey(
                    on_delete=django.db.models.deletion.CASCADE,
                    related_name='completion_requests',
                    to='core.papersettingduty',
                )),
                ('faculty', models.ForeignKey(
                    on_delete=django.db.models.deletion.CASCADE,
                    related_name='paper_setting_completion_requests',
                    to='core.faculty',
                )),
            ],
            options={
                'verbose_name': 'Paper setting completion request',
                'verbose_name_plural': 'Paper setting completion requests',
                'ordering': ['-submitted_at'],
            },
        ),
        migrations.AddConstraint(
            model_name='papersettingcompletionrequest',
            constraint=models.UniqueConstraint(
                condition=Q(status='pending'),
                fields=('duty', 'faculty'),
                name='uniq_pending_paper_set_completion_per_duty_faculty',
            ),
        ),
    ]
