from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('hostel', '0030_formcategory_targeting'),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.AddField(model_name='visitor', name='email',
            field=models.EmailField(blank=True, max_length=254)),
        migrations.AddField(model_name='visitor', name='purpose_type',
            field=models.CharField(choices=[('admission','Admission'),('guardian','Guardian / Relative'),('guest','Guest')],
                                   default='guardian', max_length=15)),
        migrations.AddField(model_name='visitor', name='id_upload',
            field=models.FileField(blank=True, null=True, upload_to='visitor_ids/')),
        migrations.AddField(model_name='visitor', name='student_name_text',
            field=models.CharField(blank=True, max_length=200)),
        migrations.AddField(model_name='visitor', name='is_college_student',
            field=models.BooleanField(default=False)),
        migrations.AddField(model_name='visitor', name='college_id_upload',
            field=models.FileField(blank=True, null=True, upload_to='college_ids/')),
        migrations.AddField(model_name='visitor', name='student_notified',
            field=models.BooleanField(default=False)),
        migrations.AddField(model_name='visitor', name='student_acknowledged',
            field=models.BooleanField(default=False)),
        migrations.AddField(model_name='visitor', name='student_acknowledged_at',
            field=models.DateTimeField(blank=True, null=True)),
        migrations.AlterField(model_name='visitor', name='purpose',
            field=models.TextField(blank=True)),
        migrations.AlterField(model_name='visitor', name='relation',
            field=models.CharField(blank=True, choices=[
                ('father','Father'),('mother','Mother'),('sibling','Sibling'),
                ('guardian','Guardian'),('relative','Relative'),('friend','Friend'),('other','Other')],
                max_length=15)),
        migrations.AlterField(model_name='visitor', name='student',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL,
                                    related_name='visitors', to='hostel.student')),
    ]
