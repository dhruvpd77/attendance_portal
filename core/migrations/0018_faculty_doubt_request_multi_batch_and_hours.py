# Generated manually for multi-batch doubt requests + nullable legacy batch.

from django.db import migrations, models


def copy_batch_to_m2m(apps, schema_editor):
    FacultyDoubtRequest = apps.get_model('core', 'FacultyDoubtRequest')
    for dr in FacultyDoubtRequest.objects.all():
        if dr.batch_id:
            dr.batches.add(dr.batch_id)


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0017_faculty_doubt_request'),
    ]

    operations = [
        migrations.AddField(
            model_name='facultydoubtrequest',
            name='batches',
            field=models.ManyToManyField(
                blank=True,
                related_name='faculty_doubt_requests',
                to='core.batch',
            ),
        ),
        migrations.AlterField(
            model_name='facultydoubtrequest',
            name='batch',
            field=models.ForeignKey(
                blank=True,
                help_text='Deprecated: use batches. Kept for older rows.',
                null=True,
                on_delete=models.deletion.CASCADE,
                to='core.batch',
            ),
        ),
        migrations.RunPython(copy_batch_to_m2m, migrations.RunPython.noop),
    ]
