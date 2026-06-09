from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('hostel', '0027_department'),
    ]

    operations = [
        migrations.AddField(
            model_name='announcement',
            name='target_hostels',
            field=models.JSONField(
                blank=True, default=list,
                help_text='List of specific hostel UUIDs (str). Empty = all hostels.'
            ),
        ),
    ]
