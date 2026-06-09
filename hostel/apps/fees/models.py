import uuid
from django.db import models
from apps.hostel.models import Student, Room


class FeeStructure(models.Model):
    """Admin defines fee amounts per room type per semester."""

    id           = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name         = models.CharField(max_length=100)          # e.g. "Sem 1 2024-25"
    academic_year = models.CharField(max_length=10)           # e.g. "2024-25"
    semester     = models.PositiveSmallIntegerField()
    room_type    = models.CharField(
        max_length=10,
        choices=[('single','Single'),('double','Double'),('triple','Triple')],
        blank=True
    )
    hostel_fee   = models.DecimalField(max_digits=8, decimal_places=2)
    mess_fee     = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    other_fee    = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    other_fee_label = models.CharField(max_length=100, blank=True)
    due_date     = models.DateField()
    is_active    = models.BooleanField(default=True)
    created_at   = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'fee_structures'

    def __str__(self):
        return f'{self.name} ({self.academic_year} Sem {self.semester})'

    @property
    def total(self):
        return self.hostel_fee + self.mess_fee + self.other_fee


class FeeRecord(models.Model):
    """One fee record per student per semester — tracks payment status."""

    class Status(models.TextChoices):
        PENDING = 'pending', 'Pending'
        PAID    = 'paid',    'Paid'
        PARTIAL = 'partial', 'Partially Paid'
        WAIVED  = 'waived',  'Waived'

    id            = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    student       = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='fee_records')
    fee_structure = models.ForeignKey(FeeStructure, on_delete=models.PROTECT)
    total_amount  = models.DecimalField(max_digits=8, decimal_places=2)
    amount_paid   = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    status        = models.CharField(max_length=10, choices=Status.choices, default=Status.PENDING)
    paid_on       = models.DateField(null=True, blank=True)
    payment_mode  = models.CharField(max_length=50, blank=True)  # Cash / Online / DD
    transaction_ref = models.CharField(max_length=100, blank=True)
    remarks       = models.TextField(blank=True)
    created_at    = models.DateTimeField(auto_now_add=True)
    updated_at    = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'fee_records'
        unique_together = ('student', 'fee_structure')
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.student.roll_number} — {self.fee_structure.name} — {self.status}'

    @property
    def balance(self):
        return self.total_amount - self.amount_paid
