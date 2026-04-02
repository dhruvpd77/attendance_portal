# Academic period (InstituteSemester) scopes all departments, exam phases, and exam logins.

import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


def forwards(apps, schema_editor):
    InstituteSemester = apps.get_model('core', 'InstituteSemester')
    sem, _ = InstituteSemester.objects.get_or_create(
        code='SY_EVEN_2026',
        defaults={'label': 'SY EVEN 2026', 'sort_order': 100},
    )
    Department = apps.get_model('core', 'Department')
    for d in Department.objects.filter(institute_semester__isnull=True).iterator():
        d.institute_semester_id = sem.pk
        d.save(update_fields=['institute_semester_id'])

    SupervisionExamPhase = apps.get_model('core', 'SupervisionExamPhase')
    for p in SupervisionExamPhase.objects.filter(institute_semester__isnull=True).iterator():
        if p.department_id:
            d = Department.objects.filter(pk=p.department_id).first()
            p.institute_semester_id = d.institute_semester_id if d else sem.pk
        else:
            p.institute_semester_id = sem.pk
        p.save(update_fields=['institute_semester_id'])

    PaperCheckingPhase = apps.get_model('core', 'PaperCheckingPhase')
    for p in PaperCheckingPhase.objects.filter(institute_semester__isnull=True).iterator():
        if p.department_id:
            d = Department.objects.filter(pk=p.department_id).first()
            p.institute_semester_id = d.institute_semester_id if d else sem.pk
        else:
            p.institute_semester_id = sem.pk
        p.save(update_fields=['institute_semester_id'])

    PaperSettingPhase = apps.get_model('core', 'PaperSettingPhase')
    for p in PaperSettingPhase.objects.filter(institute_semester__isnull=True).iterator():
        if p.department_id:
            d = Department.objects.filter(pk=p.department_id).first()
            p.institute_semester_id = d.institute_semester_id if d else sem.pk
        else:
            p.institute_semester_id = sem.pk
        p.save(update_fields=['institute_semester_id'])

    DepartmentExamProfile = apps.get_model('core', 'DepartmentExamProfile')
    for prof in DepartmentExamProfile.objects.filter(institute_semester__isnull=True).iterator():
        if prof.department_id:
            d = Department.objects.filter(pk=prof.department_id).first()
            prof.institute_semester_id = d.institute_semester_id if d else sem.pk
        else:
            prof.institute_semester_id = sem.pk
        prof.save(update_fields=['institute_semester_id'])


def backwards(apps, schema_editor):
    pass


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0031_paper_checking_subject_credit'),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.CreateModel(
            name='InstituteSemester',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                (
                    'code',
                    models.SlugField(
                        help_text='Stable id shown in URLs/exports, e.g. SY_EVEN_2026',
                        max_length=64,
                        unique=True,
                    ),
                ),
                ('label', models.CharField(max_length=200)),
                (
                    'sort_order',
                    models.PositiveSmallIntegerField(
                        default=0,
                        help_text='Higher sorts first in pickers.',
                    ),
                ),
                ('created_at', models.DateTimeField(auto_now_add=True)),
            ],
            options={
                'verbose_name': 'Institute semester',
                'verbose_name_plural': 'Institute semesters',
                'ordering': ['-sort_order', '-pk'],
            },
        ),
        migrations.AddField(
            model_name='department',
            name='institute_semester',
            field=models.ForeignKey(
                null=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name='departments',
                to='core.institutesemester',
            ),
        ),
        migrations.RenameField(
            model_name='department',
            old_name='semester',
            new_name='dr_export_semester_label',
        ),
        migrations.AddField(
            model_name='departmentexamprofile',
            name='institute_semester',
            field=models.ForeignKey(
                help_text='Academic period for this login; hub accounts set explicitly, others usually match linked department.',
                null=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name='exam_profiles',
                to='core.institutesemester',
            ),
        ),
        migrations.AddField(
            model_name='supervisionexamphase',
            name='institute_semester',
            field=models.ForeignKey(
                null=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name='supervision_phases',
                to='core.institutesemester',
            ),
        ),
        migrations.AddField(
            model_name='papercheckingphase',
            name='institute_semester',
            field=models.ForeignKey(
                null=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name='paper_checking_phases',
                to='core.institutesemester',
            ),
        ),
        migrations.AddField(
            model_name='papersettingphase',
            name='institute_semester',
            field=models.ForeignKey(
                null=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name='paper_setting_phases',
                to='core.institutesemester',
            ),
        ),
        migrations.RunPython(forwards, backwards),
        migrations.AlterField(
            model_name='department',
            name='institute_semester',
            field=models.ForeignKey(
                on_delete=django.db.models.deletion.CASCADE,
                related_name='departments',
                to='core.institutesemester',
            ),
        ),
        migrations.AlterField(
            model_name='departmentexamprofile',
            name='institute_semester',
            field=models.ForeignKey(
                help_text='Academic period for this login; hub accounts set explicitly, others usually match linked department.',
                on_delete=django.db.models.deletion.CASCADE,
                related_name='exam_profiles',
                to='core.institutesemester',
            ),
        ),
        migrations.AlterField(
            model_name='supervisionexamphase',
            name='institute_semester',
            field=models.ForeignKey(
                on_delete=django.db.models.deletion.CASCADE,
                related_name='supervision_phases',
                to='core.institutesemester',
            ),
        ),
        migrations.AlterField(
            model_name='papercheckingphase',
            name='institute_semester',
            field=models.ForeignKey(
                on_delete=django.db.models.deletion.CASCADE,
                related_name='paper_checking_phases',
                to='core.institutesemester',
            ),
        ),
        migrations.AlterField(
            model_name='papersettingphase',
            name='institute_semester',
            field=models.ForeignKey(
                on_delete=django.db.models.deletion.CASCADE,
                related_name='paper_setting_phases',
                to='core.institutesemester',
            ),
        ),
        migrations.RemoveConstraint(
            model_name='supervisionexamphase',
            name='uniq_supervision_phase_per_hub_name',
        ),
        migrations.AddConstraint(
            model_name='supervisionexamphase',
            constraint=models.UniqueConstraint(
                condition=models.Q(hub_coordinator__isnull=False, department__isnull=True),
                fields=('hub_coordinator', 'name', 'institute_semester'),
                name='uniq_supervision_phase_per_hub_name',
            ),
        ),
        migrations.RemoveConstraint(
            model_name='papercheckingphase',
            name='uniq_paper_check_phase_hub_name',
        ),
        migrations.RemoveConstraint(
            model_name='papercheckingphase',
            name='uniq_paper_check_phase_institute_name',
        ),
        migrations.AddConstraint(
            model_name='papercheckingphase',
            constraint=models.UniqueConstraint(
                condition=models.Q(hub_coordinator__isnull=False, institute_scope=False),
                fields=('hub_coordinator', 'name', 'institute_semester'),
                name='uniq_paper_check_phase_hub_name',
            ),
        ),
        migrations.AddConstraint(
            model_name='papercheckingphase',
            constraint=models.UniqueConstraint(
                condition=models.Q(institute_scope=True),
                fields=('institute_semester', 'name'),
                name='uniq_paper_check_phase_institute_name',
            ),
        ),
        migrations.RemoveConstraint(
            model_name='papersettingphase',
            name='uniq_paper_set_phase_hub_name',
        ),
        migrations.RemoveConstraint(
            model_name='papersettingphase',
            name='uniq_paper_set_phase_institute_name',
        ),
        migrations.AddConstraint(
            model_name='papersettingphase',
            constraint=models.UniqueConstraint(
                condition=models.Q(hub_coordinator__isnull=False, institute_scope=False),
                fields=('hub_coordinator', 'name', 'institute_semester'),
                name='uniq_paper_set_phase_hub_name',
            ),
        ),
        migrations.AddConstraint(
            model_name='papersettingphase',
            constraint=models.UniqueConstraint(
                condition=models.Q(institute_scope=True),
                fields=('institute_semester', 'name'),
                name='uniq_paper_set_phase_institute_name',
            ),
        ),
        migrations.RemoveConstraint(
            model_name='departmentexamprofile',
            name='uniq_exam_subunit_per_parent',
        ),
        migrations.AddConstraint(
            model_name='departmentexamprofile',
            constraint=models.UniqueConstraint(
                condition=models.Q(parent__isnull=False),
                fields=('parent', 'subunit_code', 'institute_semester'),
                name='uniq_exam_subunit_per_parent',
            ),
        ),
    ]
