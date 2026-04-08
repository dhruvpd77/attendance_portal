# Generated manually

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0051_risk_criteria_and_faculty_sidebar_toggles'),
    ]

    operations = [
        migrations.AddField(
            model_name='facultydoubtrequest',
            name='location',
            field=models.CharField(
                blank=True,
                help_text='Where the doubt session takes place (room, lab, online link, etc.).',
                max_length=255,
            ),
        ),
    ]
