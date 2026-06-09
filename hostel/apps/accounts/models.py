import uuid
from django.contrib.auth.models import AbstractBaseUser, BaseUserManager, PermissionsMixin
from django.db import models


class UserManager(BaseUserManager):
    def create_user(self, email, password=None, **extra):
        if not email:
            raise ValueError('Email is required')
        user = self.model(email=self.normalize_email(email), **extra)
        user.set_password(password)
        user.save(using=self._db)
        return user

    def create_superuser(self, email, password, **extra):
        extra.setdefault('role', User.Role.SUPER_ADMIN)
        extra.setdefault('is_staff', True)
        extra.setdefault('is_superuser', True)
        return self.create_user(email, password, **extra)


class User(AbstractBaseUser, PermissionsMixin):
    class Role(models.TextChoices):
        SUPER_ADMIN = 'superadmin',   'Admin 1'
        ADMIN       = 'admin',        'Admin 2'
        WARDEN      = 'warden',       'Warden'
        HOD         = 'hod',          'HOD'
        STUDENT     = 'student',      'Student'
        SECURITY    = 'security',     'Security Guard'
        MAINTENANCE = 'maintenance',  'Maintenance Incharge'
        MESS        = 'mess',         'Mess Management'

    id         = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    email      = models.EmailField(unique=True)
    name       = models.CharField(max_length=150)
    role       = models.CharField(max_length=40, choices=Role.choices, default=Role.STUDENT)
    phone      = models.CharField(max_length=15, blank=True)
    department = models.CharField(max_length=150, blank=True, help_text='Department name — required for HOD role')
    is_active = models.BooleanField(default=True)
    is_staff  = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    objects = UserManager()
    USERNAME_FIELD  = 'email'
    REQUIRED_FIELDS = ['name']

    class Meta:
        db_table = 'users'

    def __str__(self):
        return f'{self.name} ({self.role})'

    @property
    def is_super_admin(self):
        return self.role == self.Role.SUPER_ADMIN

    @property
    def is_admin_or_warden(self):
        return self.role in (self.Role.SUPER_ADMIN, self.Role.ADMIN, self.Role.WARDEN)

    @property
    def is_security(self):
        return self.role == self.Role.SECURITY

    @property
    def can_manage_gate_pass(self):
        return self.role in (self.Role.SUPER_ADMIN, self.Role.ADMIN,
                             self.Role.WARDEN, self.Role.SECURITY)

    @property
    def is_custom_staff_role(self):
        """True for any dynamically-created staff role not in the built-in set."""
        built_in = {
            self.Role.SUPER_ADMIN, self.Role.ADMIN, self.Role.WARDEN,
            self.Role.HOD, self.Role.STUDENT, self.Role.SECURITY,
            self.Role.MAINTENANCE, self.Role.MESS,
        }
        return self.role not in built_in
