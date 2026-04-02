# Institute-wide exam credit rules (nullable department) + paper_checking task.

import django.db.models.deletion
from django.db import migrations, models
from django.db.models import Q


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0034_paper_setting_completion_and_credit_rules'),
    ]

    operations = [
        migrations.RemoveConstraint(
            model_name='departmentexamcreditrule',
            name='uniq_dept_exam_credit_rule',
        ),
        migrations.AlterField(
            model_name='departmentexamcreditrule',
            name='department',
            field=models.ForeignKey(
                blank=True,
                help_text='Empty = institute default applied to all departments (unless a department has its own rule).',
                null=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name='exam_credit_rules',
                to='core.department',
            ),
        ),
        migrations.AlterField(
            model_name='departmentexamcreditrule',
            name='task',
            field=models.CharField(
                choices=[
                    ('paper_setting', 'Paper setting (per approved duty)'),
                    ('supervision', 'Supervision (per completed duty)'),
                    ('paper_checking', 'Paper checking (theory fallback: credit per paper if no subject rule)'),
                ],
                max_length=24,
            ),
        ),
        migrations.AddConstraint(
            model_name='departmentexamcreditrule',
            constraint=models.UniqueConstraint(
                condition=Q(department__isnull=False),
                fields=('department', 'task', 'phase_bucket', 'subject_name'),
                name='uniq_dept_exam_credit_rule',
            ),
        ),
        migrations.AddConstraint(
            model_name='departmentexamcreditrule',
            constraint=models.UniqueConstraint(
                condition=Q(department__isnull=True),
                fields=('task', 'phase_bucket', 'subject_name'),
                name='uniq_institute_exam_credit_rule',
            ),
        ),
    ]
