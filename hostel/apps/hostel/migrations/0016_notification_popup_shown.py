from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('hostel', '0015_add_leave_application'),
    ]

    operations = [
        migrations.AddField(
            model_name='notification',
            name='popup_shown',
            field=models.BooleanField(default=True),
        ),
    ]
