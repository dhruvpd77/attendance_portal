# Generated migration for PhaseHoliday

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0002_lecture_adjustment'),
    ]

    operations = [
        migrations.CreateModel(
            name='PhaseHoliday',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('phase', models.CharField(max_length=10)),
                ('date', models.DateField()),
                ('department', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='core.department')),
            ],
            options={
                'ordering': ['phase', 'date'],
                'unique_together': {('department', 'phase', 'date')},
            },
        ),
    ]
