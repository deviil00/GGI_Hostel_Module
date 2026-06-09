from django.contrib import admin
from apps.fees.models import FeeStructure, FeeRecord


@admin.register(FeeStructure)
class FeeStructureAdmin(admin.ModelAdmin):
    list_display  = ['name', 'academic_year', 'semester', 'room_type', 'hostel_fee', 'mess_fee', 'other_fee', 'due_date', 'is_active']
    list_filter   = ['academic_year', 'is_active']
    list_editable = ['is_active']


@admin.register(FeeRecord)
class FeeRecordAdmin(admin.ModelAdmin):
    list_display  = ['student', 'fee_structure', 'total_amount', 'amount_paid', 'status', 'paid_on']
    list_filter   = ['status', 'fee_structure__academic_year']
    search_fields = ['student__roll_number', 'student__name']
    list_editable = ['status']
    raw_id_fields = ['student']
