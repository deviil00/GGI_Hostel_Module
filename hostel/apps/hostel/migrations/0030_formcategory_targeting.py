from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('hostel', '0029_form_categories'),
    ]

    operations = [
        migrations.AddField(
            model_name='formcategory',
            name='target_years',
            field=models.JSONField(blank=True, default=list,
                                   help_text='List of year ints, e.g. [1,2]. Empty = all years.'),
        ),
        migrations.AddField(
            model_name='formcategory',
            name='target_hostels',
            field=models.JSONField(blank=True, default=list,
                                   help_text='List of hostel UUID strings. Empty = all hostels.'),
        ),
        migrations.AddField(
            model_name='formcategory',
            name='target_depts',
            field=models.JSONField(blank=True, default=list,
                                   help_text='List of department name strings. Empty = all depts.'),
        ),
    ]
