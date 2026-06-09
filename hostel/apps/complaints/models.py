import uuid
from django.db import models
from django.utils import timezone
from apps.hostel.models import Student, Room
from apps.accounts.models import User


class Complaint(models.Model):
    class Category(models.TextChoices):
        ELECTRICAL     = 'electrical',     'Electrical'
        PLUMBING       = 'plumbing',       'Plumbing'
        FURNITURE      = 'furniture',      'Furniture'
        INTERNET       = 'internet',       'Internet / WiFi'
        CLEANLINESS    = 'cleanliness',    'Cleanliness'
        CARPENTRY      = 'carpentry',      'Carpentry'
        INFRASTRUCTURE = 'infrastructure', 'Infrastructure'
        CIVIL          = 'civil',          'Civil Work'
        PAINTING       = 'painting',       'Painting / Whitewash'
        PEST_CONTROL   = 'pest_control',   'Pest Control'
        AC             = 'ac',             'AC / HVAC'
        SECURITY_INFRA = 'security_infra', 'Security Infrastructure'
        OTHER          = 'other',          'Other'

    class Status(models.TextChoices):
        SUBMITTED   = 'submitted',   'Submitted'
        VERIFIED    = 'verified',    'Verified by Warden'
        FORWARDED   = 'forwarded',   'Forwarded to Maintenance'
        IN_PROGRESS = 'in_progress', 'In Progress'
        RESOLVED    = 'resolved',    'Resolved'
        CLOSED      = 'closed',      'Closed'

    class Priority(models.TextChoices):
        LOW    = 'low',    'Low'
        MEDIUM = 'medium', 'Medium'
        HIGH   = 'high',   'High'
        URGENT = 'urgent', 'Urgent'

    id              = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    # Reporter — either a student OR a staff member (warden / superadmin)
    student         = models.ForeignKey(Student, on_delete=models.CASCADE,
                                        related_name='complaints', null=True, blank=True)
    raised_by_staff = models.ForeignKey(User, on_delete=models.SET_NULL,
                                        null=True, blank=True,
                                        related_name='staff_complaints')
    room            = models.ForeignKey(Room, on_delete=models.SET_NULL, null=True, blank=True)
    location        = models.CharField(max_length=300, blank=True,
                                       help_text="Area/location for staff-reported complaints")
    category        = models.CharField(max_length=20, choices=Category.choices)
    title       = models.CharField(max_length=200)
    description = models.TextField()

    # Priority is set by warden after verification, not by student
    priority    = models.CharField(max_length=8, choices=Priority.choices,
                                   null=True, blank=True)
    status      = models.CharField(max_length=15, choices=Status.choices,
                                   default=Status.SUBMITTED)

    # Warden fields
    warden_remarks  = models.TextField(blank=True)
    assigned_to     = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True,
        related_name='assigned_complaints',
        limit_choices_to={'role': User.Role.MAINTENANCE}
    )
    forwarded_at    = models.DateTimeField(null=True, blank=True)

    # Maintenance fields
    maintenance_remarks = models.TextField(blank=True)

    # Legacy — kept for backward compat
    remarks     = models.TextField(blank=True)

    resolved_at = models.DateTimeField(null=True, blank=True)
    closed_at   = models.DateTimeField(null=True, blank=True)
    created_at  = models.DateTimeField(auto_now_add=True)
    updated_at  = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'complaints'
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.title} — {self.status}'

    @property
    def is_staff_complaint(self):
        return self.student_id is None

    @property
    def reporter_name(self):
        if self.student_id:
            return self.student.name
        if self.raised_by_staff_id:
            return self.raised_by_staff.name or self.raised_by_staff.email
        return 'Unknown'

    @property
    def reporter_role(self):
        if self.student_id:
            return f'Student — {self.student.roll_number}'
        if self.raised_by_staff_id:
            return self.raised_by_staff.role.replace('_', ' ').title()
        return ''

    @property
    def location_display(self):
        if self.room:
            hostel = self.room.hostel.name if self.room.hostel_id else ''
            return f'{self.room.room_number}{"  —  " + hostel if hostel else ""}'
        return self.location or '—'

    @property
    def is_overdue(self):
        """True if forwarded to maintenance but not resolved within 24 hours."""
        if self.forwarded_at and self.status in (
            self.Status.FORWARDED, self.Status.IN_PROGRESS
        ):
            return (timezone.now() - self.forwarded_at).total_seconds() > 86400
        return False

    @property
    def hours_since_forwarded(self):
        if self.forwarded_at:
            return round((timezone.now() - self.forwarded_at).total_seconds() / 3600, 1)
        return None


class ComplaintTimeline(models.Model):
    """Immutable audit trail for every action taken on a complaint."""
    id         = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    complaint  = models.ForeignKey(Complaint, on_delete=models.CASCADE,
                                   related_name='timeline')
    status     = models.CharField(max_length=15, choices=Complaint.Status.choices,
                                  blank=True)
    comment    = models.TextField(blank=True)
    updated_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                   related_name='complaint_actions')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'complaint_timeline'
        ordering = ['created_at']

    def __str__(self):
        return f'{self.complaint_id} → {self.status} by {self.updated_by}'
