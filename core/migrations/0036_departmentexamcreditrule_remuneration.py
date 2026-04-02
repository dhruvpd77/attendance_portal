# Generated manually for exam remuneration (₹ per same unit as credit).

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0035_exam_credit_rule_institute_scope'),
    ]

    operations = [
        migrations.AddField(
            model_name='departmentexamcreditrule',
            name='remuneration',
            field=models.DecimalField(
                decimal_places=2,
                default=0,
                help_text='Remuneration (₹) per same unit as credit: per duty (setting/supervision) or per paper (checking fallback).',
                max_digits=12,
            ),
        ),
    ]
