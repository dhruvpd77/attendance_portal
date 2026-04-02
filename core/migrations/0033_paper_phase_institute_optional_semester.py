# Generated manually — exam-section phases are not tied to one institute period.

from django.db import migrations, models
import django.db.models.deletion


def _dedupe_institute_phase_names(apps, model_name: str):
    InstituteSemester = apps.get_model('core', 'InstituteSemester')
    Model = apps.get_model('core', model_name)
    from django.db.models import Count

    dup_names = (
        Model.objects.filter(institute_scope=True)
        .values('name')
        .annotate(n=Count('id'))
        .filter(n__gt=1)
    )
    for row in dup_names:
        name = row['name']
        phases = list(Model.objects.filter(institute_scope=True, name=name).order_by('pk'))
        for p in phases[1:]:
            code = ''
            if p.institute_semester_id:
                sem = InstituteSemester.objects.filter(pk=p.institute_semester_id).first()
                code = (sem.code if sem else '') or str(p.institute_semester_id)
            suffix = f' ({code})' if code else f' ({p.pk})'
            new_name = (name + suffix)[:80]
            p.name = new_name
            p.save(update_fields=['name'])


def institute_phases_clear_semester(apps, schema_editor):
    for model_name in ('PaperCheckingPhase', 'PaperSettingPhase'):
        _dedupe_institute_phase_names(apps, model_name)
        Model = apps.get_model('core', model_name)
        Model.objects.filter(institute_scope=True).update(institute_semester=None)


def noop_reverse(apps, schema_editor):
    pass


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0032_institute_semester'),
    ]

    operations = [
        migrations.RemoveConstraint(
            model_name='papercheckingphase',
            name='uniq_paper_check_phase_institute_name',
        ),
        migrations.RemoveConstraint(
            model_name='papersettingphase',
            name='uniq_paper_set_phase_institute_name',
        ),
        migrations.AlterField(
            model_name='papercheckingphase',
            name='institute_semester',
            field=models.ForeignKey(
                blank=True,
                help_text='Empty for exam-section (institute) phases — shared across all institute periods. '
                'Required for department / hub phases.',
                null=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name='paper_checking_phases',
                to='core.institutesemester',
            ),
        ),
        migrations.AlterField(
            model_name='papersettingphase',
            name='institute_semester',
            field=models.ForeignKey(
                blank=True,
                help_text='Empty for exam-section (institute) phases — shared across all institute periods. '
                'Required for department / hub phases.',
                null=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name='paper_setting_phases',
                to='core.institutesemester',
            ),
        ),
        migrations.RunPython(institute_phases_clear_semester, noop_reverse),
        migrations.AddConstraint(
            model_name='papercheckingphase',
            constraint=models.UniqueConstraint(
                condition=models.Q(institute_scope=True),
                fields=('name',),
                name='uniq_paper_check_phase_institute_name',
            ),
        ),
        migrations.AddConstraint(
            model_name='papercheckingphase',
            constraint=models.CheckConstraint(
                condition=models.Q(institute_scope=True)
                | models.Q(institute_semester__isnull=False),
                name='paper_check_phase_semester_if_dept_or_hub',
            ),
        ),
        migrations.AddConstraint(
            model_name='papersettingphase',
            constraint=models.UniqueConstraint(
                condition=models.Q(institute_scope=True),
                fields=('name',),
                name='uniq_paper_set_phase_institute_name',
            ),
        ),
        migrations.AddConstraint(
            model_name='papersettingphase',
            constraint=models.CheckConstraint(
                condition=models.Q(institute_scope=True)
                | models.Q(institute_semester__isnull=False),
                name='paper_set_phase_semester_if_dept_or_hub',
            ),
        ),
    ]
