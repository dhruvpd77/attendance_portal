# Generated for Result section - Exam phases and marks

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0011_extralecture'),
    ]

    operations = [
        migrations.CreateModel(
            name='ExamPhase',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=50)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('department', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='core.department')),
            ],
            options={
                'verbose_name': 'Exam phase',
                'verbose_name_plural': 'Exam phases',
                'ordering': ['department', 'name'],
                'unique_together': {('department', 'name')},
            },
        ),
        migrations.CreateModel(
            name='ExamPhaseSubject',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('exam_phase', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='core.examphase')),
                ('subject', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='core.subject')),
            ],
            options={
                'verbose_name': 'Exam phase subject',
                'verbose_name_plural': 'Exam phase subjects',
                'ordering': ['exam_phase', 'subject'],
                'unique_together': {('exam_phase', 'subject')},
            },
        ),
        migrations.CreateModel(
            name='StudentMark',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('marks_obtained', models.DecimalField(blank=True, decimal_places=2, max_digits=6, null=True)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('updated_at', models.DateTimeField(auto_now=True)),
                ('exam_phase', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='core.examphase')),
                ('student', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='core.student')),
                ('subject', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='core.subject')),
            ],
            options={
                'verbose_name': 'Student mark',
                'verbose_name_plural': 'Student marks',
                'ordering': ['exam_phase', 'subject', 'student'],
                'unique_together': {('student', 'exam_phase', 'subject')},
            },
        ),
    ]
