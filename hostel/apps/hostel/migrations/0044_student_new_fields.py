import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('hostel', '0043_gatepasscategory_max_hours'),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.AddField(
            model_name='student',
            name='state',
            field=models.CharField(blank=True, max_length=100),
        ),
        migrations.AddField(
            model_name='student',
            name='country',
            field=models.CharField(blank=True, default='India', max_length=100),
        ),
        migrations.AddField(
            model_name='student',
            name='type_of_entry',
            field=models.CharField(
                blank=True,
                choices=[('incampus', 'InCampus'), ('outcampus', 'OutCampus')],
                max_length=10,
            ),
        ),
        migrations.AddField(
            model_name='student',
            name='reporting_date',
            field=models.DateField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='student',
            name='is_active',
            field=models.BooleanField(default=True),
        ),
        migrations.AddField(
            model_name='student',
            name='inactive_reason',
            field=models.CharField(
                blank=True,
                choices=[('left_hostel', 'Left Hostel'), ('left_college', 'Left College')],
                max_length=20,
            ),
        ),
        migrations.AddField(
            model_name='student',
            name='inactive_remarks',
            field=models.TextField(blank=True),
        ),
        migrations.AddField(
            model_name='student',
            name='exit_date',
            field=models.DateField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='student',
            name='exit_time',
            field=models.TimeField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='student',
            name='marked_inactive_by',
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name='student_deactivations',
                to=settings.AUTH_USER_MODEL,
            ),
        ),
        migrations.AddField(
            model_name='student',
            name='marked_inactive_at',
            field=models.DateTimeField(blank=True, null=True),
        ),
    ]
