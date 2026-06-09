from django.contrib import admin
from apps.complaints.models import Complaint


@admin.register(Complaint)
class ComplaintAdmin(admin.ModelAdmin):
    list_display  = ['student', 'category', 'title', 'priority', 'status', 'created_at']
    list_filter   = ['status', 'priority', 'category']
    search_fields = ['student__roll_number', 'student__name', 'title']
    list_editable = ['status']
    raw_id_fields = ['student']
