from django.core.management.base import BaseCommand
from apps.accounts.models import User


class Command(BaseCommand):
    help = 'Create default Admin 1 superadmin account if it does not exist'

    def handle(self, *args, **kwargs):
        if User.objects.filter(email='admin1@hostel.com').exists():
            self.stdout.write('Admin 1 account already exists — skipping.')
            return
        u = User.objects.create_user(email='admin1@hostel.com', password='Admin@1234')
        u.role = 'superadmin'
        u.name = 'Admin 1'
        u.is_staff = True
        u.is_superuser = True
        u.save()
        self.stdout.write(self.style.SUCCESS('Admin 1 created. email=admin1@hostel.com password=Admin@1234'))
