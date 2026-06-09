from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion
import uuid


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('hostel', '0026_announcement_custom_audiences'),
    ]

    operations = [
        migrations.CreateModel(
            name='Department',
            fields=[
                ('id', models.UUIDField(default=uuid.uuid4, editable=False, primary_key=True, serialize=False)),
                ('name', models.CharField(help_text='Full name e.g. Computer Science & Engineering', max_length=150, unique=True)),
                ('code', models.CharField(blank=True, help_text='Short code e.g. CSE, ECE, ME', max_length=20)),
                ('is_active', models.BooleanField(default=True)),
                ('created_by', models.ForeignKey(
                    null=True, on_delete=django.db.models.deletion.SET_NULL,
                    related_name='created_departments', to=settings.AUTH_USER_MODEL
                )),
                ('created_at', models.DateTimeField(auto_now_add=True)),
            ],
            options={'db_table': 'departments', 'ordering': ['name']},
        ),
    ]
