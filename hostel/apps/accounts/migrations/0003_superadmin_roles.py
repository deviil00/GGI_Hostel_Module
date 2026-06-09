from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('accounts', '0002_qr_gate_pass_security_role'),
    ]

    operations = [
        migrations.AlterField(
            model_name='user',
            name='role',
            field=models.CharField(
                choices=[
                    ('superadmin',   'Admin 1'),
                    ('admin',        'Admin 2'),
                    ('warden',       'Warden'),
                    ('student',      'Student'),
                    ('security',     'Security Guard'),
                    ('maintenance',  'Maintenance Incharge'),
                    ('mess',         'Mess Management'),
                ],
                default='student',
                max_length=15,
            ),
        ),
    ]
