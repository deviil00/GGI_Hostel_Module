import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('hostel', '0016_notification_popup_shown'),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.CreateModel(
            name='RoleModulePermission',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False)),
                ('role', models.CharField(max_length=15, unique=True)),
                ('allowed_modules', models.JSONField(default=list)),
                ('updated_at', models.DateTimeField(auto_now=True)),
                ('updated_by', models.ForeignKey(
                    blank=True, null=True,
                    on_delete=django.db.models.deletion.SET_NULL,
                    related_name='permission_updates',
                    to=settings.AUTH_USER_MODEL,
                )),
            ],
            options={'db_table': 'role_module_permissions'},
        ),
    ]
