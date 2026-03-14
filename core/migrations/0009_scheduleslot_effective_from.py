# Schedule versioning: effective_from date for timetable changes

from datetime import date
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0008_lecturecancellation'),
    ]

    operations = [
        migrations.AddField(
            model_name='scheduleslot',
            name='effective_from',
            field=models.DateField(default=date(2000, 1, 1), help_text='Date from which this slot applies. Original schedule uses 2000-01-01.'),
            preserve_default=True,
        ),
        migrations.AlterUniqueTogether(
            name='scheduleslot',
            unique_together={('department', 'batch', 'day', 'time_slot', 'effective_from')},
        ),
    ]
