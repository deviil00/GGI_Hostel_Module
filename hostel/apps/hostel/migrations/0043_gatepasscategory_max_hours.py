from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('hostel', '0042_add_ids_and_quota_policy'),
    ]

    operations = [
        migrations.RenameField(
            model_name='gatepasscategory',
            old_name='max_days',
            new_name='max_hours',
        ),
    ]
