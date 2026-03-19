# Generated for Add Extra Lecture feature

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0010_add_student_phone_parents_contact'),
    ]

    operations = [
        migrations.CreateModel(
            name='ExtraLecture',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('date', models.DateField()),
                ('time_slot', models.CharField(max_length=50)),
                ('room_number', models.CharField(blank=True, max_length=50)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('batch', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='core.batch')),
                ('faculty', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='core.faculty')),
                ('subject', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='core.subject')),
            ],
            options={
                'verbose_name': 'Extra lecture',
                'verbose_name_plural': 'Extra lectures',
                'ordering': ['-date', 'batch', 'time_slot'],
                'unique_together': {('date', 'batch', 'time_slot')},
            },
        ),
    ]
