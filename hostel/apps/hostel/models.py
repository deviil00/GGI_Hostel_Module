import uuid
from django.db import models
from apps.accounts.models import User


# ── Room category system ─────────────────────────────────────────────────────
# Each entry: (code, human-readable label, default capacity)
ROOM_CATEGORIES = [
    ('4S+CWR',     '4-Seater + Common Washroom',         4),
    ('4S+CWR+AC',  '4-Seater + Common Washroom + AC',    4),
    ('4S+WR',      '4-Seater + Attached Washroom',       4),
    ('4S+WR+AC',   '4-Seater + Attached Washroom + AC',  4),
    ('3S+CWR',     '3-Seater + Common Washroom',         3),
    ('3S+CWR+AC',  '3-Seater + Common Washroom + AC',    3),
    ('3S+WR',      '3-Seater + Attached Washroom',       3),
    ('3S+WR+AC',   '3-Seater + Attached Washroom + AC',  3),
    ('2S+CWR',     '2-Seater + Common Washroom',         2),
    ('2S+CWR+AC',  '2-Seater + Common Washroom + AC',    2),
    ('2S+WR',      '2-Seater + Attached Washroom',       2),
    ('2S+WR+AC',   '2-Seater + Attached Washroom + AC',  2),
    ('1S+CWR',     'Single + Common Washroom',           1),
    ('1S+CWR+AC',  'Single + Common Washroom + AC',      1),
    ('1S+WR',      'Single + Attached Washroom',         1),
    ('1S+WR+AC',   'Single + Attached Washroom + AC',    1),
    # Common / amenity rooms — no beds (capacity 0)
    ('GYM',            'Gym',                0),
    ('SPORTS ROOM',    'Sports Room',         0),
    ('ENTERTAINMENT',  'Entertainment Room',  0),
    ('MESS',           'Mess Hall',           0),
    ('STUDY ROOM',     'Study Room',          0),
    ('SITTING ARENA',  'Sitting Arena',       0),
    ('STORE ROOM',     'Store Room',          0),
]

# Code → default capacity lookup
ROOM_CATEGORY_CAPACITY = {code: cap for code, _, cap in ROOM_CATEGORIES}

# Grouped for template optgroups
ROOM_CATEGORIES_GROUPED = [
    ('Student Rooms', [(c, l, p) for c, l, p in ROOM_CATEGORIES if p > 0]),
    ('Common / Amenity', [(c, l, p) for c, l, p in ROOM_CATEGORIES if p == 0]),
]

# Code → label lookup for display
ROOM_CATEGORY_LABELS = {code: label for code, label, _ in ROOM_CATEGORIES}
# ─────────────────────────────────────────────────────────────────────────────


class Hostel(models.Model):
    class Type(models.TextChoices):
        BOYS  = 'boys',  'Boys'
        GIRLS = 'girls', 'Girls'
        STAFF = 'staff', 'Staff'

    id      = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name    = models.CharField(max_length=200, unique=True)
    type    = models.CharField(max_length=5, choices=Type.choices)
    address = models.TextField(blank=True)
    warden  = models.ForeignKey(
        User, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='managed_hostels',
        limit_choices_to={'role': User.Role.WARDEN}
    )
    mess_incharge = models.ForeignKey(
        User, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='mess_hostels',
        limit_choices_to={'role': User.Role.MESS}
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'hostels'

    def __str__(self):
        return self.name


class Room(models.Model):
    class Status(models.TextChoices):
        VACANT      = 'vacant',      'Vacant'
        OCCUPIED    = 'occupied',    'Occupied'
        RESERVED    = 'reserved',    'Reserved'
        MAINTENANCE = 'maintenance', 'Under Maintenance'

    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    hostel      = models.ForeignKey(Hostel, on_delete=models.CASCADE, related_name='rooms')
    room_number = models.CharField(max_length=20)
    floor       = models.PositiveSmallIntegerField(default=1)
    wing        = models.CharField(max_length=50, blank=True, help_text="e.g. A-Wing, North Block")
    room_type   = models.CharField(max_length=100, default='Standard',
                                   help_text="Admin-defined type e.g. AC Double, Non-AC Triple, VIP")
    capacity    = models.PositiveSmallIntegerField(default=2)
    status      = models.CharField(max_length=15, choices=Status.choices, default=Status.VACANT)
    amenities   = models.ManyToManyField('RoomAmenity', blank=True, related_name='rooms')

    class Meta:
        db_table = 'rooms'
        unique_together = ('hostel', 'room_number')
        ordering = ['hostel', 'floor', 'room_number']

    def __str__(self):
        return f'{self.hostel.name} — Room {self.room_number}'

    @property
    def current_occupants(self):
        return self.allocations.filter(status='active').count()

    @property
    def beds_available(self):
        return self.capacity - self.current_occupants


class Student(models.Model):
    class Year(models.TextChoices):
        FIRST  = '1', '1st Year'
        SECOND = '2', '2nd Year'
        THIRD  = '3', '3rd Year'
        FOURTH = '4', '4th Year'

    class Gender(models.TextChoices):
        MALE   = 'male',   'Male'
        FEMALE = 'female', 'Female'
        OTHER  = 'other',  'Other'

    class FlagColor(models.TextChoices):
        GREEN  = 'green',  'Green'
        YELLOW = 'yellow', 'Yellow'
        RED    = 'red',    'Red'

    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    user        = models.OneToOneField(
        User, on_delete=models.CASCADE, related_name='student_profile',
        null=True, blank=True
    )
    roll_number    = models.CharField(max_length=50, unique=True, db_index=True)
    name           = models.CharField(max_length=150)
    email          = models.EmailField(blank=True)
    phone          = models.CharField(max_length=15, blank=True)
    department     = models.CharField(max_length=100)
    year           = models.CharField(max_length=1, choices=Year.choices)
    semester       = models.PositiveSmallIntegerField(default=1)
    is_resident    = models.BooleanField(default=True)
    # ── extended profile fields ──
    gender         = models.CharField(max_length=6, choices=Gender.choices, blank=True)
    date_of_birth  = models.DateField(null=True, blank=True)
    guardian_name  = models.CharField(max_length=150, blank=True)
    guardian_phone = models.CharField(max_length=15, blank=True)
    address        = models.TextField(blank=True)
    photo          = models.ImageField(upload_to='student_photos/', blank=True, null=True)
    # ── documents ──
    aadhar_number  = models.CharField(max_length=12, blank=True)
    aadhar_doc     = models.FileField(upload_to='student_docs/aadhar/', blank=True, null=True)
    college_id_number = models.CharField(max_length=50, blank=True)
    college_id_doc = models.FileField(upload_to='student_docs/college_id/', blank=True, null=True)
    # ── guardian extended ──
    guardian_relation = models.CharField(max_length=50, blank=True)
    guardian_email    = models.EmailField(blank=True)
    # ── flag system ──
    flag_color   = models.CharField(max_length=10, choices=FlagColor.choices, default=FlagColor.GREEN)
    flag_note    = models.TextField(blank=True)
    flag_set_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True,
                                     related_name='student_flags')
    flag_set_at  = models.DateTimeField(null=True, blank=True)
    # ── profile edit permission ──
    can_edit_profile = models.BooleanField(default=False)
    created_at     = models.DateTimeField(auto_now_add=True)
    updated_at     = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'students'
        ordering = ['roll_number']

    def __str__(self):
        return f'{self.roll_number} — {self.name}'

    @property
    def current_room(self):
        alloc = self.allocations.filter(status='active').select_related('room__hostel').first()
        return alloc.room if alloc else None


class Semester(models.Model):
    class SemType(models.TextChoices):
        ODD  = 'odd',  'Odd Semester'
        EVEN = 'even', 'Even Semester'

    id         = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name       = models.CharField(max_length=100)          # e.g. "2024-25 Odd"
    sem_type   = models.CharField(max_length=4, choices=SemType.choices)
    start_date = models.DateField()
    end_date   = models.DateField(null=True, blank=True)
    is_current = models.BooleanField(default=False)
    closed_at  = models.DateTimeField(null=True, blank=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='semesters_created')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'semesters'
        ordering = ['-start_date']

    def __str__(self):
        return self.name

    def save(self, *args, **kwargs):
        # Ensure only one semester is marked current at a time
        if self.is_current:
            Semester.objects.exclude(pk=self.pk).update(is_current=False)
        super().save(*args, **kwargs)

    @classmethod
    def current(cls):
        return cls.objects.filter(is_current=True).first()


class RoomAllocation(models.Model):
    class Status(models.TextChoices):
        ACTIVE   = 'active',   'Active'
        CHECKOUT = 'checkout', 'Checked Out'

    id           = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    student      = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='allocations')
    room         = models.ForeignKey(Room, on_delete=models.CASCADE, related_name='allocations')
    bed          = models.ForeignKey('Bed', on_delete=models.SET_NULL, null=True, blank=True,
                                     related_name='allocations')
    semester     = models.ForeignKey(Semester, on_delete=models.SET_NULL, null=True, blank=True,
                                     related_name='allocations')
    allocated_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    check_in     = models.DateField()
    check_out    = models.DateField(null=True, blank=True)
    status       = models.CharField(max_length=10, choices=Status.choices, default=Status.ACTIVE)
    notes        = models.TextField(blank=True)
    created_at   = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'room_allocations'
        ordering = ['-created_at']
        # one active allocation per student at a time
        constraints = [
            models.UniqueConstraint(
                fields=['student'],
                condition=models.Q(status='active'),
                name='one_active_allocation_per_student'
            )
        ]

    def __str__(self):
        return f'{self.student.roll_number} → {self.room}'


class Attendance(models.Model):
    class Status(models.TextChoices):
        PRESENT = 'present', 'Present'
        ABSENT  = 'absent',  'Absent'
        LEAVE   = 'leave',   'On Leave'

    id        = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    student   = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='attendance_records')
    date      = models.DateField()
    status    = models.CharField(max_length=10, choices=Status.choices, default=Status.PRESENT)
    marked_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    remarks   = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table      = 'attendance'
        unique_together = ('student', 'date')
        ordering      = ['-date', 'student__roll_number']

    def __str__(self):
        return f'{self.student.roll_number} — {self.date} — {self.status}'


# ─── Room Amenities ───────────────────────────────────────────────────────────

class RoomAmenity(models.Model):
    """Admin-defined room feature/amenity (e.g. AC, WiFi, Hot Water)."""
    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name        = models.CharField(max_length=100, unique=True)
    icon        = models.CharField(max_length=60, default='bi-star', blank=True,
                                   help_text="Bootstrap Icons class, e.g. bi-snow2")
    description = models.CharField(max_length=300, blank=True)
    created_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                    related_name='created_amenities')
    created_at  = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'room_amenities'
        ordering = ['name']

    def __str__(self):
        return self.name


class RoomAmenityRecord(models.Model):
    """Tracks a specific amenity installed in a specific room — with condition and date."""

    class Condition(models.TextChoices):
        GOOD    = 'good',    'Good'
        FAIR    = 'fair',    'Fair'
        POOR    = 'poor',    'Poor'
        DAMAGED = 'damaged', 'Damaged'
        MISSING = 'missing', 'Missing'

    id        = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    room      = models.ForeignKey(Room,        on_delete=models.CASCADE, related_name='amenity_records')
    amenity   = models.ForeignKey(RoomAmenity, on_delete=models.CASCADE, related_name='records')
    condition = models.CharField(max_length=10, choices=Condition.choices, default=Condition.GOOD)
    notes     = models.TextField(blank=True)
    added_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                  related_name='added_amenity_records')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'room_amenity_records'
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.amenity.name} — {self.room}'


# ─── Master Data ──────────────────────────────────────────────────────────────

class Department(models.Model):
    """Admin-managed master list of courses / departments in the institution."""
    id         = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name       = models.CharField(max_length=150, unique=True,
                                  help_text='Full name e.g. Computer Science & Engineering')
    code       = models.CharField(max_length=20, blank=True,
                                  help_text='Short code e.g. CSE, ECE, ME')
    is_active  = models.BooleanField(default=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                   related_name='created_departments')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'departments'
        ordering = ['name']

    def __str__(self):
        return f'{self.code} — {self.name}' if self.code else self.name


# ─── Announcements ────────────────────────────────────────────────────────────

class AnnouncementAudience(models.Model):
    """Admin-defined custom audience group (e.g. NSS Volunteers, Sports Team, Block A)."""
    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name        = models.CharField(max_length=100, unique=True)
    description = models.CharField(max_length=200, blank=True,
                                   help_text='Brief note on who this group covers')
    icon        = models.CharField(max_length=60, default='bi-people-fill', blank=True)
    color       = models.CharField(max_length=7, default='#059669',
                                   help_text='Hex colour e.g. #059669')
    created_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                    related_name='custom_ann_audiences')
    created_at  = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'announcement_audiences'
        ordering = ['name']

    def __str__(self):
        return self.name


class AnnouncementCategory(models.Model):
    """Admin-defined announcement category (e.g. Academics, Gate Pass, Events)."""
    id         = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name       = models.CharField(max_length=100, unique=True)
    icon       = models.CharField(max_length=60, default='bi-megaphone', blank=True)
    color      = models.CharField(max_length=7, default='#1565C0',
                                  help_text='Hex colour e.g. #1565C0')
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                   related_name='ann_categories')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'announcement_categories'
        ordering = ['name']

    def __str__(self):
        return self.name


class Announcement(models.Model):
    class Priority(models.TextChoices):
        INFO    = 'info',    'Info'
        WARNING = 'warning', 'Warning'
        URGENT  = 'urgent',  'Urgent'

    class Audience(models.TextChoices):
        ALL         = 'all',         'Everyone'
        STUDENTS    = 'students',    'Students Only'
        STAFF_ALL   = 'staff_all',   'All Staff'
        WARDEN      = 'warden',      'Wardens'
        SECURITY    = 'security',    'Security Guards'
        MAINTENANCE = 'maintenance', 'Maintenance Staff'
        MESS        = 'mess',        'Mess Staff'
        CUSTOM      = 'custom',      'Custom Group'

    GENDER_CHOICES = [('all','All'), ('boys','Boys Hostel'), ('girls','Girls Hostel')]

    id            = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    title         = models.CharField(max_length=200)
    content       = models.TextField()
    category      = models.ForeignKey(AnnouncementCategory, on_delete=models.SET_NULL,
                                      null=True, blank=True, related_name='announcements')
    priority      = models.CharField(max_length=10, choices=Priority.choices,
                                     default=Priority.INFO)
    audience      = models.CharField(max_length=20, choices=Audience.choices,
                                     default=Audience.ALL)
    custom_audience = models.ForeignKey(
        AnnouncementAudience, on_delete=models.SET_NULL, null=True, blank=True,
        related_name='announcements',
        help_text='Set only when audience = custom'
    )
    target_years  = models.JSONField(default=list, blank=True,
                                     help_text='List of year values e.g. ["1","2"]')
    target_depts   = models.JSONField(default=list, blank=True,
                                      help_text='List of department names')
    target_gender  = models.CharField(max_length=6, choices=GENDER_CHOICES, default='all')
    target_hostels = models.JSONField(default=list, blank=True,
                                      help_text='List of specific hostel UUIDs (str). Empty = all hostels.')
    attachment    = models.FileField(upload_to='announcements/attachments/',
                                     blank=True, null=True)
    is_active     = models.BooleanField(default=True)
    expires_at    = models.DateField(null=True, blank=True)
    created_by    = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                      related_name='announcements')
    created_at    = models.DateTimeField(auto_now_add=True)
    updated_at    = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'announcements'
        ordering = ['-created_at']

    def __str__(self):
        return self.title


# ─── Registration Form System ─────────────────────────────────────────────────

class FormCategory(models.Model):
    """Defines who receives a form — by year, hostel, and/or department."""
    id             = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name           = models.CharField(max_length=100, unique=True)
    icon           = models.CharField(max_length=50, default='bi-file-earmark-text')
    color          = models.CharField(max_length=7, default='#1565C0')
    # Audience targeting (empty list = all)
    target_years   = models.JSONField(default=list, blank=True,
                                      help_text='List of year ints, e.g. [1,2]. Empty = all years.')
    target_hostels = models.JSONField(default=list, blank=True,
                                      help_text='List of hostel UUID strings. Empty = all hostels.')
    target_depts   = models.JSONField(default=list, blank=True,
                                      help_text='List of department name strings. Empty = all depts.')
    created_by     = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                       related_name='created_form_categories')
    created_at     = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'form_categories'
        ordering = ['name']

    def __str__(self):
        return self.name


class RegistrationForm(models.Model):
    """Admin-defined registration form template for room allocation."""
    id            = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    title         = models.CharField(max_length=200)
    description   = models.TextField(blank=True)
    category      = models.CharField(max_length=100, blank=True,
                                     help_text="Legacy text category — use form_category FK instead")
    form_category = models.ForeignKey(FormCategory, on_delete=models.SET_NULL,
                                      null=True, blank=True, related_name='forms')
    is_active   = models.BooleanField(default=True)
    deadline    = models.DateField(null=True, blank=True)
    max_preferences = models.PositiveSmallIntegerField(default=3,
                      help_text="How many room preferences the student can specify")
    created_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                    related_name='created_reg_forms')
    created_at  = models.DateTimeField(auto_now_add=True)
    updated_at  = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'registration_forms'
        ordering = ['-created_at']

    def __str__(self):
        return self.title


class FormField(models.Model):
    """A single dynamic field inside a RegistrationForm."""
    class FieldType(models.TextChoices):
        TEXT     = 'text',     'Short Text'
        TEXTAREA = 'textarea', 'Long Text'
        NUMBER   = 'number',   'Number'
        EMAIL    = 'email',    'Email'
        DATE     = 'date',     'Date'
        SELECT   = 'select',   'Dropdown'
        RADIO    = 'radio',    'Radio Buttons'
        CHECKBOX = 'checkbox', 'Checkboxes'

    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    form        = models.ForeignKey(RegistrationForm, on_delete=models.CASCADE,
                                    related_name='fields')
    label       = models.CharField(max_length=200)
    field_type  = models.CharField(max_length=10, choices=FieldType.choices,
                                   default=FieldType.TEXT)
    options     = models.JSONField(default=list, blank=True,
                                   help_text="Options list for select/radio/checkbox fields")
    is_required = models.BooleanField(default=True)
    order       = models.PositiveSmallIntegerField(default=0)
    help_text   = models.CharField(max_length=300, blank=True)

    class Meta:
        db_table = 'form_fields'
        ordering = ['order']

    def __str__(self):
        return f'{self.form.title} — {self.label}'


class RegistrationSubmission(models.Model):
    """A student's submitted application to a RegistrationForm."""
    class Status(models.TextChoices):
        SUBMITTED    = 'submitted',    'Submitted'
        UNDER_REVIEW = 'under_review', 'Under Review'
        APPROVED     = 'approved',     'Approved'
        REJECTED     = 'rejected',     'Rejected'
        WAITLISTED   = 'waitlisted',   'Waitlisted'

    id           = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    form         = models.ForeignKey(RegistrationForm, on_delete=models.CASCADE,
                                     related_name='submissions')
    student      = models.ForeignKey(Student, on_delete=models.CASCADE,
                                     related_name='reg_submissions')
    status       = models.CharField(max_length=15, choices=Status.choices,
                                    default=Status.SUBMITTED)
    submitted_at = models.DateTimeField(auto_now_add=True)
    reviewed_by  = models.ForeignKey(User, on_delete=models.SET_NULL,
                                     null=True, blank=True,
                                     related_name='reviewed_submissions')
    reviewed_at  = models.DateTimeField(null=True, blank=True)
    admin_notes  = models.TextField(blank=True)
    allocated_room = models.ForeignKey(Room, on_delete=models.SET_NULL,
                                       null=True, blank=True,
                                       related_name='reg_allocations')

    class Meta:
        db_table = 'registration_submissions'
        ordering = ['-submitted_at']
        unique_together = ('form', 'student')

    def __str__(self):
        return f'{self.student.roll_number} → {self.form.title}'


class SubmissionResponse(models.Model):
    """One field's answer within a RegistrationSubmission."""
    id         = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    submission = models.ForeignKey(RegistrationSubmission, on_delete=models.CASCADE,
                                   related_name='responses')
    field      = models.ForeignKey(FormField, on_delete=models.CASCADE)
    value      = models.TextField(blank=True)

    class Meta:
        db_table = 'submission_responses'
        unique_together = ('submission', 'field')

    def __str__(self):
        return f'{self.submission} — {self.field.label}'


# ─── Gate Pass ───────────────────────────────────────────────────────────────

class GatePass(models.Model):
    class PassType(models.TextChoices):
        DAY_OUTING = 'day_outing', 'Day Outing'
        OVERNIGHT  = 'overnight',  'Overnight Leave'
        EMERGENCY  = 'emergency',  'Emergency Leave'

    class Status(models.TextChoices):
        PENDING  = 'pending',  'Pending'
        APPROVED = 'approved', 'Approved'
        REJECTED = 'rejected', 'Rejected'
        RETURNED = 'returned', 'Returned'

    id                   = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    student              = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='gate_passes')
    category             = models.ForeignKey('GatePassCategory', on_delete=models.SET_NULL,
                                             null=True, blank=True, related_name='gate_passes')
    pass_type            = models.CharField(max_length=15, choices=PassType.choices,
                                            default=PassType.DAY_OUTING)
    reason               = models.TextField()
    destination          = models.CharField(max_length=300)
    departure_date       = models.DateField()
    departure_time       = models.TimeField()
    expected_return_date = models.DateField()
    expected_return_time = models.TimeField()
    actual_return_date   = models.DateField(null=True, blank=True)
    actual_return_time   = models.TimeField(null=True, blank=True)
    status               = models.CharField(max_length=15, choices=Status.choices, default=Status.PENDING)
    approved_by          = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True,
                                             related_name='approved_gate_passes')
    approved_at          = models.DateTimeField(null=True, blank=True)
    admin_remarks        = models.TextField(blank=True)
    is_overstayed        = models.BooleanField(default=False)

    # ── Guard permission (guard must permit before student can generate QR) ──
    guard_exit_permitted    = models.BooleanField(default=False)
    guard_exit_permitted_at = models.DateTimeField(null=True, blank=True)
    guard_exit_permitted_by = models.ForeignKey(
        User, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='exit_permitted_passes'
    )
    guard_entry_permitted    = models.BooleanField(default=False)
    guard_entry_permitted_at = models.DateTimeField(null=True, blank=True)
    guard_entry_permitted_by = models.ForeignKey(
        User, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='entry_permitted_passes'
    )

    # ── QR / Movement tracking ──
    # Exit QR — generated only after guard permits (valid 5 min)
    exit_qr_token        = models.CharField(max_length=600, blank=True)
    exit_qr_generated_at = models.DateTimeField(null=True, blank=True)
    exit_qr_scanned      = models.BooleanField(default=False)
    exit_time            = models.DateTimeField(null=True, blank=True)
    exit_allowed_by      = models.ForeignKey(
        User, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='exit_allowed_passes'
    )

    # Entry QR — generated only after guard permits (valid 5 min)
    entry_qr_token        = models.CharField(max_length=600, blank=True)
    entry_qr_generated_at = models.DateTimeField(null=True, blank=True)
    entry_qr_scanned      = models.BooleanField(default=False)
    entry_time            = models.DateTimeField(null=True, blank=True)
    entry_allowed_by      = models.ForeignKey(
        User, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='entry_allowed_passes'
    )

    # Human-readable reference ID: GP-YYYY-NNNNN
    gp_id                = models.CharField(max_length=20, unique=True, blank=True, db_index=True)

    created_at           = models.DateTimeField(auto_now_add=True)
    updated_at           = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'gate_passes'
        ordering = ['-created_at']

    def save(self, *args, **kwargs):
        if not self.gp_id:
            from django.utils import timezone as _tz
            year = _tz.now().year
            last = GatePass.objects.filter(gp_id__startswith=f'GP-{year}-').order_by('-gp_id').first()
            seq = 1
            if last and last.gp_id:
                try:
                    seq = int(last.gp_id.split('-')[-1]) + 1
                except (ValueError, IndexError):
                    pass
            self.gp_id = f'GP-{year}-{seq:05d}'
        super().save(*args, **kwargs)

    def __str__(self):
        return f'{self.gp_id} — {self.student.roll_number} ({self.status})'

    def check_overstay(self):
        """Flag as overstayed if expected return has passed and not yet returned."""
        from datetime import datetime as dt
        from django.utils import timezone as tz
        if self.status == 'approved':
            expected_dt = dt.combine(self.expected_return_date, self.expected_return_time)
            expected_aware = tz.make_aware(expected_dt)
            if tz.now() > expected_aware:
                self.is_overstayed = True
                self.save(update_fields=['is_overstayed'])
        return self.is_overstayed


class RoomPreference(models.Model):
    """Ordered room preferences given by a student in a submission."""
    id                = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    submission        = models.ForeignKey(RegistrationSubmission, on_delete=models.CASCADE,
                                          related_name='preferences')
    preference_number = models.PositiveSmallIntegerField()
    hostel            = models.ForeignKey(Hostel, on_delete=models.SET_NULL,
                                          null=True, blank=True)
    room_type         = models.CharField(max_length=100, blank=True)
    floor               = models.PositiveSmallIntegerField(null=True, blank=True)
    preferred_amenities = models.JSONField(default=list, blank=True,
                          help_text="List of RoomAmenity PKs the student prefers")

    class Meta:
        db_table = 'room_preferences'
        unique_together = ('submission', 'preference_number')
        ordering = ['preference_number']

    def __str__(self):
        return f'{self.submission} — Pref {self.preference_number}'


# ─── Bed (Bed-level allocation) ───────────────────────────────────────────────

class Bed(models.Model):
    class Status(models.TextChoices):
        AVAILABLE   = 'available',   'Available'
        OCCUPIED    = 'occupied',    'Occupied'
        RESERVED    = 'reserved',    'Reserved'
        MAINTENANCE = 'maintenance', 'Under Maintenance'

    id         = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    room       = models.ForeignKey(Room, on_delete=models.CASCADE, related_name='beds')
    bed_number = models.CharField(max_length=10, help_text="e.g. A, B, 1, 2")
    status     = models.CharField(max_length=15, choices=Status.choices, default=Status.AVAILABLE)

    class Meta:
        db_table = 'beds'
        unique_together = ('room', 'bed_number')
        ordering = ['room', 'bed_number']

    def __str__(self):
        return f'{self.room} — Bed {self.bed_number}'


# ─── In-App Notifications ─────────────────────────────────────────────────────

class Notification(models.Model):
    class Type(models.TextChoices):
        ANNOUNCEMENT = 'announcement', 'Announcement'
        GATE_PASS    = 'gate_pass',    'Gate Pass'
        COMPLAINT    = 'complaint',    'Complaint'
        VISITOR      = 'visitor',      'Visitor'
        DISCIPLINE   = 'discipline',   'Discipline'
        MESS         = 'mess',         'Mess'
        GENERAL      = 'general',      'General'

    id                = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    recipient         = models.ForeignKey(User, on_delete=models.CASCADE, related_name='notifications')
    title             = models.CharField(max_length=200)
    message           = models.TextField()
    notification_type = models.CharField(max_length=15, choices=Type.choices, default=Type.GENERAL)
    link              = models.CharField(max_length=500, blank=True)
    is_read           = models.BooleanField(default=False)
    popup_shown       = models.BooleanField(default=True)   # False = needs to pop up
    created_at        = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'notifications'
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.recipient.email} — {self.title}'


# ─── Visitor Management ───────────────────────────────────────────────────────

class Visitor(models.Model):
    class Status(models.TextChoices):
        PENDING     = 'pending',     'Pending Approval'
        APPROVED    = 'approved',    'Approved'
        REJECTED    = 'rejected',    'Rejected'
        CHECKED_IN  = 'checked_in',  'Checked In'
        CHECKED_OUT = 'checked_out', 'Checked Out'

    class Relation(models.TextChoices):
        FATHER   = 'father',   'Father'
        MOTHER   = 'mother',   'Mother'
        SIBLING  = 'sibling',  'Sibling'
        GUARDIAN = 'guardian', 'Guardian'
        RELATIVE = 'relative', 'Relative'
        FRIEND   = 'friend',   'Friend'
        OTHER    = 'other',    'Other'

    class PurposeType(models.TextChoices):
        ADMISSION = 'admission', 'Admission'
        GUARDIAN  = 'guardian',  'Guardian / Relative'
        GUEST     = 'guest',     'Guest'

    # Core
    id             = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    hostel         = models.ForeignKey(Hostel, on_delete=models.SET_NULL, related_name='visitors', null=True, blank=True)
    visitor_name   = models.CharField(max_length=150)
    email          = models.EmailField(blank=True)
    phone          = models.CharField(max_length=15)
    id_proof_type  = models.CharField(max_length=50, default='Aadhar')
    id_number      = models.CharField(max_length=50)
    id_upload      = models.FileField(upload_to='visitor_ids/', null=True, blank=True)
    visit_date     = models.DateField()
    expected_in_time  = models.TimeField()
    expected_out_time = models.TimeField()
    actual_in_time    = models.DateTimeField(null=True, blank=True)
    actual_out_time   = models.DateTimeField(null=True, blank=True)
    # Purpose
    purpose_type   = models.CharField(max_length=15, choices=PurposeType.choices,
                                      default=PurposeType.GUARDIAN)
    purpose        = models.TextField(blank=True)          # legacy / guest reason
    # Guardian-specific
    student        = models.ForeignKey(Student, on_delete=models.SET_NULL,
                                       related_name='visitors', null=True, blank=True)
    student_name_text = models.CharField(max_length=200, blank=True)
    relation       = models.CharField(max_length=15, choices=Relation.choices, blank=True)
    # Guest-specific
    is_college_student    = models.BooleanField(default=False)
    college_id_upload     = models.FileField(upload_to='college_ids/', null=True, blank=True)
    # Status & approval
    status         = models.CharField(max_length=15, choices=Status.choices, default=Status.PENDING)
    approved_by    = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True,
                                        related_name='approved_visitors')
    remarks        = models.TextField(blank=True)
    # Student notification
    student_notified     = models.BooleanField(default=False)
    student_acknowledged = models.BooleanField(default=False)
    student_acknowledged_at = models.DateTimeField(null=True, blank=True)
    created_at     = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'visitors'
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.visitor_name} → {self.hostel.name} ({self.visit_date})'


# ─── Mess Management ──────────────────────────────────────────────────────────

class MessMenu(models.Model):
    class DayOfWeek(models.TextChoices):
        MONDAY    = 'monday',    'Monday'
        TUESDAY   = 'tuesday',   'Tuesday'
        WEDNESDAY = 'wednesday', 'Wednesday'
        THURSDAY  = 'thursday',  'Thursday'
        FRIDAY    = 'friday',    'Friday'
        SATURDAY  = 'saturday',  'Saturday'
        SUNDAY    = 'sunday',    'Sunday'

    class MealType(models.TextChoices):
        BREAKFAST = 'breakfast', 'Breakfast'
        LUNCH     = 'lunch',     'Lunch'
        SNACKS    = 'snacks',    'Evening Snacks'
        DINNER    = 'dinner',    'Dinner'

    id         = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    hostel     = models.ForeignKey(Hostel, on_delete=models.CASCADE, null=True, blank=True,
                                    related_name='menus',
                                    help_text="Leave blank to apply to all hostels")
    day        = models.CharField(max_length=15, choices=DayOfWeek.choices)
    meal_type  = models.CharField(max_length=15, choices=MealType.choices)
    items      = models.TextField(help_text="Comma-separated or bullet list of items")
    is_active  = models.BooleanField(default=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                    related_name='created_menus')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'mess_menus'
        ordering = ['day', 'meal_type']

    def __str__(self):
        hostel_name = self.hostel.name if self.hostel else 'All Hostels'
        return f'{hostel_name} — {self.day} {self.meal_type}'


class MessFeedback(models.Model):
    id         = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    student    = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='mess_feedback')
    hostel     = models.ForeignKey(Hostel, on_delete=models.CASCADE)
    date       = models.DateField()
    meal_type  = models.CharField(max_length=15, choices=MessMenu.MealType.choices)
    rating     = models.PositiveSmallIntegerField(help_text="1 (Poor) to 5 (Excellent)")
    comment    = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'mess_feedback'
        unique_together = ('student', 'date', 'meal_type')
        ordering = ['-date']

    def __str__(self):
        return f'{self.student.roll_number} — {self.date} {self.meal_type} ({self.rating}★)'


class MessDailyMenu(models.Model):
    """Date-specific menu entry — created by bulk upload or manual daily edit."""
    id             = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    hostel         = models.ForeignKey(Hostel, on_delete=models.CASCADE, null=True, blank=True,
                                        related_name='daily_menus',
                                        help_text="Null = applies to all hostels")
    menu_date      = models.DateField()
    meal_type      = models.CharField(max_length=15, choices=MessMenu.MealType.choices)
    items          = models.TextField()
    original_items = models.TextField(blank=True,
                                       help_text="Items as originally uploaded; untouched by edits")
    is_overridden  = models.BooleanField(default=False,
                                          help_text="True when manually edited after upload")
    upload_month   = models.CharField(max_length=7, blank=True,
                                       help_text="YYYY-MM of the upload batch this came from")
    created_by     = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                        related_name='created_daily_menus')
    updated_by     = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                        related_name='updated_daily_menus')
    created_at     = models.DateTimeField(auto_now_add=True)
    updated_at     = models.DateTimeField(auto_now=True)

    class Meta:
        db_table        = 'mess_daily_menus'
        unique_together = ('hostel', 'menu_date', 'meal_type')
        ordering        = ['menu_date', 'meal_type']

    def __str__(self):
        h = self.hostel.name if self.hostel else 'All Hostels'
        return f'{h} — {self.menu_date} {self.meal_type}'


class MessWastageRecord(models.Model):
    """Daily food wastage entry filled by Mess Incharge."""
    id         = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    hostel     = models.ForeignKey(Hostel, on_delete=models.CASCADE,
                                    related_name='wastage_records')
    date       = models.DateField()
    meal_type  = models.CharField(max_length=15, choices=MessMenu.MealType.choices)
    wastage_kg = models.DecimalField(max_digits=6, decimal_places=2,
                                      help_text="Wastage quantity in kilograms")
    notes      = models.TextField(blank=True)
    filled_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                    related_name='wastage_filled')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table        = 'mess_wastage_records'
        unique_together = ('hostel', 'date', 'meal_type')
        ordering        = ['-date', 'meal_type']

    def __str__(self):
        return f'{self.hostel.name} — {self.date} {self.meal_type}: {self.wastage_kg} kg'


class MessItemWastage(models.Model):
    """Per-item daily wastage entry: how much was prepared, consumed, and wasted."""
    UNIT_KG    = 'kg'
    UNIT_LITRE = 'litre'
    UNIT_PCS   = 'pcs'
    UNIT_CHOICES = [(UNIT_KG, 'kg'), (UNIT_LITRE, 'litre'), (UNIT_PCS, 'pcs')]

    id           = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    hostel       = models.ForeignKey(Hostel, on_delete=models.CASCADE, related_name='item_wastages',
                                      null=True, blank=True)
    date         = models.DateField()
    meal_type    = models.CharField(max_length=15, choices=MessMenu.MealType.choices)
    item_name    = models.CharField(max_length=200)
    prepared_qty = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    consumed_qty = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    wasted_qty   = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    unit         = models.CharField(max_length=10, choices=UNIT_CHOICES, default=UNIT_KG)
    filled_by    = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                      related_name='item_wastages_filled')
    notes        = models.TextField(blank=True)
    created_at   = models.DateTimeField(auto_now_add=True)
    updated_at   = models.DateTimeField(auto_now=True)

    class Meta:
        db_table        = 'mess_item_wastages'
        unique_together = ('hostel', 'date', 'meal_type', 'item_name')
        ordering        = ['-date', 'meal_type', 'item_name']

    def __str__(self):
        hostel_name = self.hostel.name if self.hostel_id else 'Campus'
        return f'{hostel_name} — {self.date} {self.meal_type}: {self.item_name}'

    @property
    def waste_pct(self):
        if self.prepared_qty:
            return round(float(self.wasted_qty) / float(self.prepared_qty) * 100, 1)
        return 0.0


# ─── Discipline Module ────────────────────────────────────────────────────────

class DisciplineCategory(models.Model):
    class Severity(models.TextChoices):
        LOW      = 'low',      'Low'
        MEDIUM   = 'medium',   'Medium'
        HIGH     = 'high',     'High'
        CRITICAL = 'critical', 'Critical'

    id                  = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name                = models.CharField(max_length=150, unique=True)
    description         = models.TextField(blank=True)
    severity            = models.CharField(max_length=10, choices=Severity.choices, default=Severity.MEDIUM)
    default_warning     = models.CharField(max_length=10,
                            choices=[('none','None'),('verbal','Verbal Warning'),
                                     ('written','Written Warning'),('final','Final Warning')],
                            default='verbal')
    fine_mandatory      = models.BooleanField(default=False)
    notify_parent_auto  = models.BooleanField(default=False,
                            help_text='Automatically suggest parent notification for this category')
    is_active           = models.BooleanField(default=True)
    created_by          = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                             related_name='discipline_categories')
    created_at          = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'discipline_categories'
        ordering = ['name']

    def __str__(self):
        return self.name

    @property
    def severity_color(self):
        return {'low':'#059669','medium':'#D97706','high':'#DC2626','critical':'#7C3AED'}.get(self.severity,'#94A3B8')


class DisciplineRecord(models.Model):
    class Status(models.TextChoices):
        PENDING             = 'pending',              'Pending'
        UNDER_INVESTIGATION = 'under_investigation',  'Under Investigation'
        RESOLVED            = 'resolved',             'Resolved'
        CLOSED              = 'closed',               'Closed'

    class WarningType(models.TextChoices):
        NONE    = 'none',    'No Warning'
        VERBAL  = 'verbal',  'Verbal Warning'
        WRITTEN = 'written', 'Written Warning'
        FINAL   = 'final',   'Final Warning'

    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    student     = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='discipline_records')
    category    = models.ForeignKey(DisciplineCategory, on_delete=models.SET_NULL,
                                     null=True, blank=True, related_name='records')

    # Incident details
    incident_date = models.DateField(default='2025-01-01')
    incident_time = models.TimeField(null=True, blank=True)
    description   = models.TextField()

    # Hostel/room snapshot at time of incident
    hostel_name   = models.CharField(max_length=200, blank=True)
    room_number   = models.CharField(max_length=50, blank=True)
    floor         = models.CharField(max_length=50, blank=True)

    # Warning
    warning_type  = models.CharField(max_length=10, choices=WarningType.choices, default=WarningType.NONE)

    # Fine
    fine_amount   = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    fine_due_date = models.DateField(null=True, blank=True)
    fine_paid     = models.BooleanField(default=False)
    fine_paid_date = models.DateField(null=True, blank=True)
    fine_paid_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True,
                                       related_name='discipline_fines_collected')

    # Status & workflow
    status        = models.CharField(max_length=25, choices=Status.choices, default=Status.PENDING)

    # Parent notification
    notify_parent       = models.BooleanField(default=False)
    parent_notified_at  = models.DateTimeField(null=True, blank=True)
    parent_notified_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True,
                                             related_name='discipline_parent_notifications')

    # Resolution
    resolution_notes = models.TextField(blank=True)
    resolved_by      = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True,
                                          related_name='discipline_resolved')
    resolved_at      = models.DateTimeField(null=True, blank=True)

    recorded_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                      related_name='discipline_records')
    created_at   = models.DateTimeField(auto_now_add=True)
    updated_at   = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'discipline_records'
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.student.roll_number} — {self.category}'

    @property
    def status_color(self):
        return {
            'pending':             '#D97706',
            'under_investigation': '#1565C0',
            'resolved':            '#059669',
            'closed':              '#64748B',
        }.get(self.status, '#94A3B8')

    @property
    def is_fine_overdue(self):
        from datetime import date
        return (self.fine_amount > 0 and not self.fine_paid
                and self.fine_due_date and self.fine_due_date < date.today())


class DisciplineEvidence(models.Model):
    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    record      = models.ForeignKey(DisciplineRecord, on_delete=models.CASCADE, related_name='evidence')
    file        = models.FileField(upload_to='discipline/evidence/')
    original_name = models.CharField(max_length=255, blank=True)
    file_type   = models.CharField(max_length=10, choices=[('image','Image'),('document','Document')],
                                    default='document')
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    uploaded_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'discipline_evidence'

    @property
    def is_image(self):
        return self.file_type == 'image'


class DisciplineAction(models.Model):
    class ActionType(models.TextChoices):
        CREATED            = 'created',             'Record Created'
        STATUS_CHANGED     = 'status_changed',      'Status Changed'
        WARNING_ISSUED     = 'warning_issued',      'Warning Issued'
        FINE_ADDED         = 'fine_added',          'Fine Added'
        FINE_PAID          = 'fine_paid',           'Fine Marked Paid'
        EVIDENCE_ADDED     = 'evidence_added',      'Evidence Added'
        PARENT_NOTIFIED    = 'parent_notified',     'Parent Notified'
        RESTRICTION_APPLIED = 'restriction_applied','Restriction Applied'
        NOTE_ADDED         = 'note_added',          'Note Added'
        RESOLVED           = 'resolved',            'Record Resolved'

    id           = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    record       = models.ForeignKey(DisciplineRecord, on_delete=models.CASCADE, related_name='actions')
    action       = models.CharField(max_length=25, choices=ActionType.choices)
    description  = models.TextField()
    performed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    created_at   = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'discipline_actions'
        ordering = ['-created_at']

    @property
    def icon(self):
        return {
            'created':              'bi-plus-circle-fill',
            'status_changed':       'bi-arrow-repeat',
            'warning_issued':       'bi-exclamation-triangle-fill',
            'fine_added':           'bi-cash-coin',
            'fine_paid':            'bi-check-circle-fill',
            'evidence_added':       'bi-paperclip',
            'parent_notified':      'bi-telephone-fill',
            'restriction_applied':  'bi-slash-circle-fill',
            'note_added':           'bi-sticky-fill',
            'resolved':             'bi-shield-check',
        }.get(self.action, 'bi-circle')


# ─── Maintenance Logs ─────────────────────────────────────────────────────────

class MaintenanceLog(models.Model):
    class Status(models.TextChoices):
        OPEN        = 'open',        'Open'
        IN_PROGRESS = 'in_progress', 'In Progress'
        RESOLVED    = 'resolved',    'Resolved'

    class Category(models.TextChoices):
        ELECTRICAL = 'electrical', 'Electrical'
        PLUMBING   = 'plumbing',   'Plumbing'
        FURNITURE  = 'furniture',  'Furniture'
        CLEANING   = 'cleaning',   'Cleaning'
        NETWORK    = 'network',    'Network/Internet'
        OTHER      = 'other',      'Other'

    id             = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    room           = models.ForeignKey(Room, on_delete=models.CASCADE, related_name='maintenance_logs')
    category       = models.CharField(max_length=15, choices=Category.choices)
    issue          = models.TextField()
    assigned_to    = models.CharField(max_length=150, blank=True)
    status         = models.CharField(max_length=15, choices=Status.choices, default=Status.OPEN)
    scheduled_date = models.DateField(null=True, blank=True)
    resolved_at    = models.DateTimeField(null=True, blank=True)
    created_by     = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                        related_name='maintenance_logs')
    created_at     = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'maintenance_logs'
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.room} — {self.category} ({self.status})'


# ─── Room Assets (Physical Inventory) ────────────────────────────────────────

class RoomAsset(models.Model):
    class Condition(models.TextChoices):
        GOOD      = 'good',      'Good'
        FAIR      = 'fair',      'Fair'
        POOR      = 'poor',      'Poor'
        DAMAGED   = 'damaged',   'Damaged'
        MISSING   = 'missing',   'Missing'

    id         = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    room       = models.ForeignKey(Room, on_delete=models.CASCADE, related_name='assets')
    name       = models.CharField(max_length=100, help_text="e.g. Fan, Cupboard, Table, Chair")
    quantity   = models.PositiveSmallIntegerField(default=1)
    condition  = models.CharField(max_length=10, choices=Condition.choices, default=Condition.GOOD)
    asset_tag  = models.CharField(max_length=50, blank=True, help_text="Asset ID / Serial number")
    notes      = models.TextField(blank=True)
    added_by   = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='room_assets')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'room_assets'
        ordering = ['name']


# ─── Role Module Permissions (SuperAdmin controlled) ──────────────────────────

STAFF_MODULES = [
    ('students',      'Students',           'bi-people-fill'),
    ('hostels',       'Hostels & Rooms',    'bi-building-fill'),
    ('attendance',    'Attendance',         'bi-calendar2-check-fill'),
    ('complaints',    'Complaints',         'bi-chat-square-text-fill'),
    ('fees',          'Fees',               'bi-cash-stack'),
    ('gate_passes',   'Gate Passes',        'bi-pass-fill'),
    ('leave',         'Leave Applications', 'bi-calendar2-x-fill'),
    ('visitors',      'Visitors',           'bi-person-check-fill'),
    ('mess',          'Mess Menu',          'bi-cup-hot-fill'),
    ('assets',        'Assets',             'bi-box-seam-fill'),
    ('discipline',    'Discipline',         'bi-exclamation-triangle-fill'),
    ('maintenance',   'Maintenance',        'bi-tools'),
    ('announcements', 'Announcements',      'bi-megaphone-fill'),
    ('reports',          'Reports',              'bi-bar-chart-fill'),
    ('reports_download', 'Download Reports',     'bi-download'),
    ('registrations',    'Registration Forms',   'bi-file-earmark-text-fill'),
    ('outsiders',        'Outsiders',            'bi-box-arrow-up-right'),
]

STUDENT_MODULES = [
    ('complaints',    'Raise Complaint',    'bi-chat-square-text-fill'),
    ('announcements', 'Announcements',      'bi-megaphone-fill'),
    ('registrations', 'Room Application',   'bi-file-earmark-text-fill'),
    ('gate_passes',   'Gate Passes',        'bi-pass-fill'),
    ('leave',         'Leave',              'bi-calendar2-x-fill'),
    ('mess',          'Mess Menu',          'bi-cup-hot-fill'),
    ('visitors',      'My Visitors',        'bi-person-check-fill'),
]

DEFAULT_ROLE_MODULES = {
    'admin': [m[0] for m in STAFF_MODULES],
    'warden': ['students', 'hostels', 'attendance', 'complaints',
               'gate_passes', 'leave', 'visitors', 'announcements', 'reports'],
    'security': ['gate_passes', 'visitors', 'outsiders'],
    'maintenance': ['maintenance', 'complaints', 'announcements'],
    'mess': ['mess', 'announcements'],
    'student': [m[0] for m in STUDENT_MODULES],
}


class RoleModulePermission(models.Model):
    role            = models.CharField(max_length=40, unique=True)
    allowed_modules = models.JSONField(default=list)
    is_active       = models.BooleanField(default=True)
    updated_at      = models.DateTimeField(auto_now=True)
    updated_by      = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True,
        related_name='permission_updates'
    )

    class Meta:
        db_table = 'role_module_permissions'

    def __str__(self):
        return f'Permissions — {self.role}'

    @classmethod
    def get_for_role(cls, role):
        """Return the set of allowed module keys for a role."""
        if role == 'superadmin':
            return {m[0] for m in STAFF_MODULES}
        perm = cls.objects.filter(role=role).first()
        if perm:
            return set(perm.allowed_modules)
        return set(DEFAULT_ROLE_MODULES.get(role, []))


class StaffProfile(models.Model):
    GENDER_CHOICES = [('male','Male'),('female','Female'),('other','Other')]
    BLOOD_GROUPS   = [('A+','A+'),('A-','A-'),('B+','B+'),('B-','B-'),
                      ('AB+','AB+'),('AB-','AB-'),('O+','O+'),('O-','O-')]

    user              = models.OneToOneField(User, on_delete=models.CASCADE, related_name='staff_profile')
    # Personal
    employee_id       = models.CharField(max_length=30, blank=True)
    date_of_birth     = models.DateField(null=True, blank=True)
    gender            = models.CharField(max_length=10, choices=GENDER_CHOICES, blank=True)
    blood_group       = models.CharField(max_length=5, choices=BLOOD_GROUPS, blank=True)
    # Contact
    address           = models.TextField(max_length=300, blank=True)
    city              = models.CharField(max_length=60, blank=True)
    state             = models.CharField(max_length=60, blank=True)
    pincode           = models.CharField(max_length=10, blank=True)
    # Emergency contact
    emergency_name    = models.CharField(max_length=100, blank=True)
    emergency_phone   = models.CharField(max_length=15, blank=True)
    emergency_relation= models.CharField(max_length=50, blank=True)
    # Professional
    qualification     = models.CharField(max_length=120, blank=True)
    department        = models.CharField(max_length=100, blank=True)
    date_of_joining   = models.DateField(null=True, blank=True)
    experience_years  = models.PositiveSmallIntegerField(null=True, blank=True)
    # Identity
    aadhar_number     = models.CharField(max_length=12, blank=True)
    college_id_number = models.CharField(max_length=50, blank=True)
    # Bio
    bio               = models.TextField(max_length=500, blank=True)
    # Official Documents
    aadhar_doc        = models.FileField(upload_to='staff_docs/aadhar/',        blank=True, null=True)
    college_id_doc    = models.FileField(upload_to='staff_docs/college_id/',    blank=True, null=True)
    appointment_letter= models.FileField(upload_to='staff_docs/appointment/',   blank=True, null=True)
    qualification_doc = models.FileField(upload_to='staff_docs/qualification/', blank=True, null=True)

    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'staff_profiles'

    def __str__(self):
        return f'Profile — {self.user.name}'

    @property
    def completion_pct(self):
        fields = [
            self.employee_id, self.date_of_birth, self.gender, self.blood_group,
            self.address, self.city, self.state,
            self.emergency_name, self.emergency_phone,
            self.qualification, self.date_of_joining,
        ]
        filled = sum(1 for f in fields if f)
        return int(filled / len(fields) * 100)

    @property
    def completion_label(self):
        pct = self.completion_pct
        if pct == 100: return 'Complete'
        if pct >= 60:  return 'Mostly Complete'
        if pct >= 30:  return 'Partial'
        return 'Incomplete'


class CustomStaffRole(models.Model):
    name       = models.CharField(max_length=60, unique=True)
    key        = models.CharField(max_length=40, unique=True)
    created_at = models.DateTimeField(auto_now_add=True)
    created_by = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True,
        related_name='created_roles'
    )

    class Meta:
        db_table = 'custom_staff_roles'
        ordering = ['name']

    def __str__(self):
        return self.name


# ─── Leave Applications ───────────────────────────────────────────────────────

class LeaveApplication(models.Model):
    class Status(models.TextChoices):
        PENDING  = 'pending',  'Pending'
        APPROVED = 'approved', 'Approved'
        REJECTED = 'rejected', 'Rejected'

    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    # Human-readable reference ID: LV-YYYY-NNNNN
    leave_id    = models.CharField(max_length=20, unique=True, blank=True, db_index=True)
    student     = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='leave_applications')
    category    = models.ForeignKey('LeaveCategory', on_delete=models.SET_NULL,
                                    null=True, blank=True, related_name='leave_applications')
    from_date   = models.DateField()
    to_date     = models.DateField()
    reason      = models.TextField()
    status      = models.CharField(max_length=10, choices=Status.choices, default=Status.PENDING)
    admin_remarks = models.TextField(blank=True)

    # Multi-stage approval
    # HOD stage — only for working-day leaves
    hod_status      = models.CharField(max_length=10, choices=Status.choices, blank=True)
    hod_approved_by = models.ForeignKey(
        'accounts.User', null=True, blank=True, on_delete=models.SET_NULL,
        related_name='hod_approved_leaves'
    )
    hod_approved_at = models.DateTimeField(null=True, blank=True)

    # Warden stage — for non-working-day leaves, or after HOD on working days
    warden_status      = models.CharField(max_length=10, choices=Status.choices, blank=True)
    warden_approved_by = models.ForeignKey(
        'accounts.User', null=True, blank=True, on_delete=models.SET_NULL,
        related_name='warden_approved_leaves'
    )
    warden_approved_at = models.DateTimeField(null=True, blank=True)

    reviewed_by = models.ForeignKey(
        'accounts.User', null=True, blank=True,
        on_delete=models.SET_NULL, related_name='reviewed_leaves'
    )
    reviewed_at = models.DateTimeField(null=True, blank=True)
    is_working_day_leave = models.BooleanField(default=False)
    created_at  = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'leave_applications'
        ordering = ['-created_at']

    def save(self, *args, **kwargs):
        if not self.leave_id:
            from django.utils import timezone as _tz
            year = _tz.now().year
            last = LeaveApplication.objects.filter(leave_id__startswith=f'LV-{year}-').order_by('-leave_id').first()
            seq = 1
            if last and last.leave_id:
                try:
                    seq = int(last.leave_id.split('-')[-1]) + 1
                except (ValueError, IndexError):
                    pass
            self.leave_id = f'LV-{year}-{seq:05d}'
        super().save(*args, **kwargs)

    def __str__(self):
        return f'{self.leave_id} — {self.student.name} ({self.status})'


# ─── Quota Policy ─────────────────────────────────────────────────────────────

class QuotaPolicy(models.Model):
    """Admin 1 sets gate pass / leave quotas per period for all or specific students."""

    class PolicyType(models.TextChoices):
        GATE_PASS = 'gate_pass', 'Gate Pass'
        LEAVE     = 'leave',     'Leave'
        BOTH      = 'both',      'Both Gate Pass & Leave'

    class Period(models.TextChoices):
        WEEKLY  = 'weekly',  'Per Week'
        MONTHLY = 'monthly', 'Per Month'
        YEARLY  = 'yearly',  'Per Year'

    id              = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    policy_type     = models.CharField(max_length=10, choices=PolicyType.choices)
    period          = models.CharField(max_length=10, choices=Period.choices)
    limit           = models.PositiveIntegerField(help_text='Maximum allowed per period')
    applies_to_all  = models.BooleanField(default=True, help_text='Apply to all students')
    student         = models.ForeignKey(
        'Student', on_delete=models.CASCADE,
        null=True, blank=True, related_name='quota_policies',
        help_text='Leave blank to apply to all students'
    )
    is_active       = models.BooleanField(default=True)
    created_by      = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='quota_policies')
    created_at      = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'quota_policies'
        ordering = ['-created_at']
        verbose_name = 'Quota Policy'
        verbose_name_plural = 'Quota Policies'

    def __str__(self):
        scope = 'All Students' if self.applies_to_all else self.student.name
        return f'{self.get_policy_type_display()} — {self.limit}/{self.get_period_display()} ({scope})'


# ─── Gate Pass Categories (SuperAdmin Configured) ─────────────────────────────

class GatePassCategory(models.Model):
    """Admin 1 configures gate pass types, approval hierarchy and time limits."""
    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name        = models.CharField(max_length=100, unique=True)
    description = models.TextField(blank=True)

    start_time  = models.TimeField(null=True, blank=True, help_text='Earliest allowed departure time')
    end_time    = models.TimeField(null=True, blank=True, help_text='Latest expected return time')
    max_hours   = models.PositiveSmallIntegerField(default=12, help_text='Maximum allowed duration in hours')

    requires_warden_approval     = models.BooleanField(default=True)
    requires_admin_approval      = models.BooleanField(default=False)
    requires_superadmin_approval = models.BooleanField(default=False)
    is_emergency                 = models.BooleanField(default=False, help_text='Emergency pass — warden can approve directly')
    is_active                    = models.BooleanField(default=True)

    created_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='gp_categories')
    created_at  = models.DateTimeField(auto_now_add=True)
    updated_at  = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'gate_pass_categories'
        ordering = ['name']

    def __str__(self):
        return self.name

    @property
    def approval_chain(self):
        chain = []
        if self.requires_warden_approval:
            chain.append('Warden')
        if self.requires_admin_approval:
            chain.append('Admin 2')
        if self.requires_superadmin_approval:
            chain.append('Admin 1')
        return ' → '.join(chain) if chain else 'Auto-approved'


# ─── Academic Calendar ────────────────────────────────────────────────────────

class AcademicCalendarEvent(models.Model):
    class EventType(models.TextChoices):
        WORKING           = 'working',           'Working Day'
        SUNDAY            = 'sunday',            'Sunday'
        HOLIDAY           = 'holiday',           'Public Holiday'
        EMERGENCY_HOLIDAY = 'emergency_holiday', 'Emergency Holiday'
        EXAM              = 'exam',              'Exam Day'
        FESTIVAL          = 'festival',          'Festival'

    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    date        = models.DateField(unique=True)
    event_type  = models.CharField(max_length=20, choices=EventType.choices)
    title       = models.CharField(max_length=200)
    description = models.TextField(blank=True)
    marked_by   = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='calendar_events')
    created_at  = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'academic_calendar_events'
        ordering = ['-date']

    def __str__(self):
        return f'{self.date} — {self.title}'

    @property
    def is_non_working(self):
        return self.event_type in (
            self.EventType.HOLIDAY,
            self.EventType.EMERGENCY_HOLIDAY,
            self.EventType.SUNDAY,
        )

    @property
    def badge_color(self):
        return {
            'working':           '#059669',
            'sunday':            '#64748B',
            'holiday':           '#D97706',
            'emergency_holiday': '#DC2626',
            'exam':              '#7C3AED',
            'festival':          '#0891B2',
        }.get(self.event_type, '#94A3B8')


# ─── Leave Categories (SuperAdmin Configured) ─────────────────────────────────

class LeaveCategory(models.Model):
    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name        = models.CharField(max_length=100, unique=True)
    description = models.TextField(blank=True)

    max_days    = models.PositiveSmallIntegerField(default=7)
    start_time  = models.TimeField(null=True, blank=True)
    end_time    = models.TimeField(null=True, blank=True)

    requires_warden_approval     = models.BooleanField(default=True)
    requires_admin_approval      = models.BooleanField(default=False)
    requires_superadmin_approval = models.BooleanField(default=False)
    is_active                    = models.BooleanField(default=True)

    created_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='leave_categories')
    created_at  = models.DateTimeField(auto_now_add=True)
    updated_at  = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'leave_categories'
        ordering = ['name']

    def __str__(self):
        return self.name

    @property
    def approval_chain(self):
        chain = []
        if self.requires_warden_approval:
            chain.append('Warden')
        if self.requires_admin_approval:
            chain.append('Admin 2')
        if self.requires_superadmin_approval:
            chain.append('Admin 1')
        return ' → '.join(chain) if chain else 'Auto-approved'


class LeaveExtensionRequest(models.Model):
    class Status(models.TextChoices):
        PENDING  = 'pending',  'Pending'
        APPROVED = 'approved', 'Approved'
        REJECTED = 'rejected', 'Rejected'

    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    leave       = models.ForeignKey(LeaveApplication, on_delete=models.CASCADE, related_name='extensions')
    new_to_date = models.DateField()
    reason      = models.TextField()
    status      = models.CharField(max_length=10, choices=Status.choices, default=Status.PENDING)
    remarks     = models.TextField(blank=True)
    approved_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='approved_extensions')
    approved_at = models.DateTimeField(null=True, blank=True)
    created_at  = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'leave_extension_requests'
        ordering = ['-created_at']

    def __str__(self):
        return f'Extension for {self.leave.student.name} → {self.new_to_date}'


# ─── Asset Management ─────────────────────────────────────────────────────────

class AssetType(models.Model):
    id          = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name        = models.CharField(max_length=100, unique=True)
    description = models.CharField(max_length=300, blank=True)
    icon        = models.CharField(max_length=60, default='bi-box-fill', blank=True)
    created_at  = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'asset_types'
        ordering = ['name']

    def __str__(self):
        return self.name


class Asset(models.Model):
    class Condition(models.TextChoices):
        GOOD              = 'good',              'Good'
        FAIR              = 'fair',              'Fair'
        DAMAGED           = 'damaged',           'Damaged'
        UNDER_MAINTENANCE = 'under_maintenance', 'Under Maintenance'
        LOST              = 'lost',              'Lost'
        DISCARDED         = 'discarded',         'Discarded'

    class Status(models.TextChoices):
        ACTIVE      = 'active',      'Active'
        TRANSFERRED = 'transferred', 'Transferred'
        DISCARDED   = 'discarded',   'Discarded'
        LOST        = 'lost',        'Lost'

    id             = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    asset_code     = models.CharField(max_length=50, unique=True, help_text='Barcode / QR / Tag ID')
    name           = models.CharField(max_length=150)
    asset_type     = models.ForeignKey(AssetType, on_delete=models.SET_NULL, null=True, blank=True, related_name='assets')

    hostel         = models.ForeignKey(Hostel, on_delete=models.SET_NULL, null=True, blank=True, related_name='assets')
    floor          = models.PositiveSmallIntegerField(null=True, blank=True)
    room           = models.ForeignKey(Room, on_delete=models.SET_NULL, null=True, blank=True, related_name='managed_assets')

    purchase_date  = models.DateField(null=True, blank=True)
    vendor_name    = models.CharField(max_length=200, blank=True)
    vendor_contact = models.CharField(max_length=100, blank=True)
    vendor_invoice = models.CharField(max_length=100, blank=True)
    purchase_cost  = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)

    condition      = models.CharField(max_length=20, choices=Condition.choices, default=Condition.GOOD)
    status         = models.CharField(max_length=15, choices=Status.choices, default=Status.ACTIVE)
    assigned_date  = models.DateField(null=True, blank=True)
    notes          = models.TextField(blank=True)

    discard_reason = models.TextField(blank=True)
    discarded_at   = models.DateTimeField(null=True, blank=True)
    discarded_by   = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='discarded_assets')

    added_by       = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='added_assets')
    created_at     = models.DateTimeField(auto_now_add=True)
    updated_at     = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'assets'
        ordering = ['hostel__name', 'floor', 'name']

    def __str__(self):
        return f'{self.asset_code} — {self.name}'

    @property
    def location_display(self):
        parts = []
        if self.hostel:
            parts.append(self.hostel.name)
        if self.floor is not None:
            parts.append(f'Floor {self.floor}')
        if self.room:
            parts.append(f'Room {self.room.room_number}')
        return ' › '.join(parts) if parts else 'Unassigned'

    @property
    def condition_color(self):
        return {
            'good':              '#059669',
            'fair':              '#D97706',
            'damaged':           '#DC2626',
            'under_maintenance': '#7C3AED',
            'lost':              '#374151',
            'discarded':         '#94A3B8',
        }.get(self.condition, '#64748B')


class AssetTransfer(models.Model):
    class Status(models.TextChoices):
        PENDING   = 'pending',   'Pending Approval'
        APPROVED  = 'approved',  'Approved'
        REJECTED  = 'rejected',  'Rejected'
        COMPLETED = 'completed', 'Completed'

    id           = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    asset        = models.ForeignKey(Asset, on_delete=models.CASCADE, related_name='transfers')

    from_hostel  = models.ForeignKey(Hostel, on_delete=models.SET_NULL, null=True, blank=True, related_name='transfers_out')
    from_room    = models.ForeignKey(Room, on_delete=models.SET_NULL, null=True, blank=True, related_name='transfers_out')
    from_floor   = models.PositiveSmallIntegerField(null=True, blank=True)

    to_hostel    = models.ForeignKey(Hostel, on_delete=models.SET_NULL, null=True, blank=True, related_name='transfers_in')
    to_room      = models.ForeignKey(Room, on_delete=models.SET_NULL, null=True, blank=True, related_name='transfers_in')
    to_floor     = models.PositiveSmallIntegerField(null=True, blank=True)

    reason        = models.TextField()
    transfer_date = models.DateField()
    status        = models.CharField(max_length=15, choices=Status.choices, default=Status.PENDING)

    requested_by  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='requested_transfers')
    approved_by   = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='approved_transfers')
    approved_at   = models.DateTimeField(null=True, blank=True)
    remarks       = models.TextField(blank=True)
    created_at    = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'asset_transfers'
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.asset.name} → {self.to_hostel}'


class AssetLog(models.Model):
    class Action(models.TextChoices):
        CREATED           = 'created',           'Created'
        UPDATED           = 'updated',           'Updated'
        TRANSFERRED       = 'transferred',       'Transferred'
        CONDITION_CHANGED = 'condition_changed',  'Condition Changed'
        DISCARDED         = 'discarded',         'Discarded'
        LOST              = 'lost',              'Marked Lost'
        REASSIGNED        = 'reassigned',        'Reassigned'

    id           = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    asset        = models.ForeignKey(Asset, on_delete=models.CASCADE, related_name='logs')
    action       = models.CharField(max_length=20, choices=Action.choices)
    description  = models.TextField()
    performed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='asset_actions')
    created_at   = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'asset_logs'
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.asset.name} — {self.action}'


# ─── Student Restriction ───────────────────────────────────────────────────────

class StudentRestriction(models.Model):
    class RestrictionType(models.TextChoices):
        GATE_PASS = 'gate_pass', 'Gate Pass'
        LEAVE     = 'leave',     'Leave'
        BOTH      = 'both',      'Gate Pass & Leave'

    id              = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    student         = models.ForeignKey('Student', on_delete=models.CASCADE, related_name='restrictions')
    restriction_type = models.CharField(max_length=10, choices=RestrictionType.choices)
    reason          = models.TextField()
    is_active       = models.BooleanField(default=True)

    restricted_by   = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                         related_name='imposed_restrictions')
    restricted_at   = models.DateTimeField(auto_now_add=True)

    lifted_by       = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True,
                                         related_name='lifted_restrictions')
    lifted_at       = models.DateTimeField(null=True, blank=True)
    lift_reason     = models.TextField(blank=True)

    class Meta:
        db_table = 'student_restrictions'
        ordering = ['-restricted_at']

    def __str__(self):
        return f'{self.student.name} — {self.restriction_type} restricted'

    @property
    def blocks_gate_pass(self):
        return self.is_active and self.restriction_type in (
            self.RestrictionType.GATE_PASS, self.RestrictionType.BOTH)

    @property
    def blocks_leave(self):
        return self.is_active and self.restriction_type in (
            self.RestrictionType.LEAVE, self.RestrictionType.BOTH)
