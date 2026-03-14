# Generated manually for lecture cancellation feature

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0007_attendancelocksetting'),
    ]

    operations = [
        migrations.CreateModel(
            name='LectureCancellation',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('date', models.DateField()),
                ('time_slot', models.CharField(max_length=50)),
                ('batch', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='core.batch')),
            ],
            options={
                'verbose_name': 'Lecture cancellation',
                'verbose_name_plural': 'Lecture cancellations',
                'ordering': ['-date', 'batch', 'time_slot'],
                'unique_together': {('date', 'batch', 'time_slot')},
            },
        ),
    ]
