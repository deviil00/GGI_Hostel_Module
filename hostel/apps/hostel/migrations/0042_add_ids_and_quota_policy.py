from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion
import uuid


def populate_gp_ids(apps, schema_editor):
    GatePass = apps.get_model('hostel', 'GatePass')
    from datetime import datetime
    rows = GatePass.objects.filter(gp_id='').order_by('created_at')
    year = datetime.now().year
    for i, obj in enumerate(rows, start=1):
        obj.gp_id = f'GP-{year}-{i:05d}'
        obj.save(update_fields=['gp_id'])


def populate_leave_ids(apps, schema_editor):
    LeaveApplication = apps.get_model('hostel', 'LeaveApplication')
    from datetime import datetime
    rows = LeaveApplication.objects.filter(leave_id='').order_by('created_at')
    year = datetime.now().year
    for i, obj in enumerate(rows, start=1):
        obj.leave_id = f'LV-{year}-{i:05d}'
        obj.save(update_fields=['leave_id'])


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('hostel', '0041_discipline_module_v2'),
    ]

    operations = [
        # Add gp_id without unique constraint first
        migrations.AddField(
            model_name='gatepass',
            name='gp_id',
            field=models.CharField(blank=True, max_length=20, default=''),
        ),
        # Populate existing rows
        migrations.RunPython(populate_gp_ids, migrations.RunPython.noop),
        # Now add unique + index
        migrations.AlterField(
            model_name='gatepass',
            name='gp_id',
            field=models.CharField(blank=True, db_index=True, max_length=20, unique=True),
        ),

        # Same for leave_id
        migrations.AddField(
            model_name='leaveapplication',
            name='leave_id',
            field=models.CharField(blank=True, max_length=20, default=''),
        ),
        migrations.RunPython(populate_leave_ids, migrations.RunPython.noop),
        migrations.AlterField(
            model_name='leaveapplication',
            name='leave_id',
            field=models.CharField(blank=True, db_index=True, max_length=20, unique=True),
        ),

        # QuotaPolicy model
        migrations.CreateModel(
            name='QuotaPolicy',
            fields=[
                ('id', models.UUIDField(default=uuid.uuid4, editable=False, primary_key=True, serialize=False)),
                ('policy_type', models.CharField(choices=[('gate_pass', 'Gate Pass'), ('leave', 'Leave'), ('both', 'Both Gate Pass & Leave')], max_length=10)),
                ('period', models.CharField(choices=[('weekly', 'Per Week'), ('monthly', 'Per Month'), ('yearly', 'Per Year')], max_length=10)),
                ('limit', models.PositiveIntegerField(help_text='Maximum allowed per period')),
                ('applies_to_all', models.BooleanField(default=True, help_text='Apply to all students')),
                ('is_active', models.BooleanField(default=True)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('created_by', models.ForeignKey(null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='quota_policies', to=settings.AUTH_USER_MODEL)),
                ('student', models.ForeignKey(blank=True, help_text='Leave blank to apply to all students', null=True, on_delete=django.db.models.deletion.CASCADE, related_name='quota_policies', to='hostel.student')),
            ],
            options={
                'verbose_name': 'Quota Policy',
                'verbose_name_plural': 'Quota Policies',
                'db_table': 'quota_policies',
                'ordering': ['-created_at'],
            },
        ),
    ]
