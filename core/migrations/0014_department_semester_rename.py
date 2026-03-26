from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0013_add_hod_and_week_lock'),
    ]

    operations = [
        migrations.RenameField(
            model_name='department',
            old_name='code',
            new_name='semester',
        ),
    ]
