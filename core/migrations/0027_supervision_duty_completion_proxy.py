# Generated manually for supervision completion + proxy tracking

from django.db import migrations, models
import django.db.models.deletion


def backfill_original_faculty(apps, schema_editor):
    SupervisionDuty = apps.get_model('core', 'SupervisionDuty')
    for d in SupervisionDuty.objects.filter(faculty_id__isnull=False).iterator():
        if d.original_faculty_id is None:
            d.original_faculty_id = d.faculty_id
            d.save(update_fields=['original_faculty'])


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0026_departmentexamprofile_hub_flag'),
    ]

    operations = [
        migrations.AddField(
            model_name='supervisionduty',
            name='block_no',
            field=models.CharField(blank=True, max_length=40),
        ),
        migrations.AddField(
            model_name='supervisionduty',
            name='room_no',
            field=models.CharField(blank=True, max_length=40),
        ),
        migrations.AddField(
            model_name='supervisionduty',
            name='completed_at',
            field=models.DateTimeField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='supervisionduty',
            name='completion_status',
            field=models.CharField(
                choices=[('open', 'Open'), ('completed', 'Completed')],
                default='open',
                max_length=20,
            ),
        ),
        migrations.AddField(
            model_name='supervisionduty',
            name='is_proxy',
            field=models.BooleanField(
                default=False,
                help_text='True if a sub-unit coordinator reassigned this duty to another faculty.',
            ),
        ),
        migrations.AddField(
            model_name='supervisionduty',
            name='original_faculty',
            field=models.ForeignKey(
                blank=True,
                help_text='Sheet assignee; unchanged when duty is marked proxy to someone else.',
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name='supervision_duties_as_original',
                to='core.faculty',
            ),
        ),
        migrations.RunPython(backfill_original_faculty, migrations.RunPython.noop),
    ]
