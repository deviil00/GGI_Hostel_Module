from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion
import uuid


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('hostel', '0025_announcement_categories_and_targeting'),
    ]

    operations = [
        # 1. Create AnnouncementAudience table
        migrations.CreateModel(
            name='AnnouncementAudience',
            fields=[
                ('id', models.UUIDField(default=uuid.uuid4, editable=False, primary_key=True, serialize=False)),
                ('name', models.CharField(max_length=100, unique=True)),
                ('description', models.CharField(blank=True, help_text='Brief note on who this group covers', max_length=200)),
                ('icon', models.CharField(blank=True, default='bi-people-fill', max_length=60)),
                ('color', models.CharField(default='#059669', help_text='Hex colour e.g. #059669', max_length=7)),
                ('created_by', models.ForeignKey(
                    null=True, on_delete=django.db.models.deletion.SET_NULL,
                    related_name='custom_ann_audiences', to=settings.AUTH_USER_MODEL
                )),
                ('created_at', models.DateTimeField(auto_now_add=True)),
            ],
            options={'db_table': 'announcement_audiences', 'ordering': ['name']},
        ),

        # 2. Add custom_audience FK to Announcement
        migrations.AddField(
            model_name='announcement',
            name='custom_audience',
            field=models.ForeignKey(
                blank=True, null=True,
                help_text='Set only when audience = custom',
                on_delete=django.db.models.deletion.SET_NULL,
                related_name='announcements',
                to='hostel.announcementaudience',
            ),
        ),

        # 3. Widen audience field to hold the new 'custom' choice
        migrations.AlterField(
            model_name='announcement',
            name='audience',
            field=models.CharField(
                choices=[
                    ('all', 'Everyone'), ('students', 'Students Only'),
                    ('staff_all', 'All Staff'), ('warden', 'Wardens'),
                    ('security', 'Security Guards'), ('maintenance', 'Maintenance Staff'),
                    ('mess', 'Mess Staff'), ('custom', 'Custom Group'),
                ],
                default='all', max_length=20,
            ),
        ),
    ]
