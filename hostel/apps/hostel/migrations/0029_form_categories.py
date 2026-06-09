import uuid
from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('hostel', '0028_announcement_target_hostels'),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.CreateModel(
            name='FormCategory',
            fields=[
                ('id',         models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False, serialize=False)),
                ('name',       models.CharField(max_length=100, unique=True)),
                ('icon',       models.CharField(default='bi-file-earmark-text', max_length=50)),
                ('color',      models.CharField(default='#1565C0', max_length=7)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('created_by', models.ForeignKey(null=True, on_delete=django.db.models.deletion.SET_NULL,
                                                 related_name='created_form_categories',
                                                 to=settings.AUTH_USER_MODEL)),
            ],
            options={'db_table': 'form_categories', 'ordering': ['name']},
        ),
        migrations.AlterField(
            model_name='registrationform',
            name='category',
            field=models.CharField(blank=True, max_length=100,
                                   help_text='Legacy text category — use form_category FK instead'),
        ),
        migrations.AddField(
            model_name='registrationform',
            name='form_category',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL,
                                    related_name='forms', to='hostel.FormCategory'),
        ),
    ]
