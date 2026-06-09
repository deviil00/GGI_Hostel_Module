from django.contrib import admin
from apps.hostel.models import Hostel, Room, Student, RoomAllocation


@admin.register(Hostel)
class HostelAdmin(admin.ModelAdmin):
    list_display  = ['name', 'type', 'warden']
    list_filter   = ['type']
    search_fields = ['name']


@admin.register(Room)
class RoomAdmin(admin.ModelAdmin):
    list_display  = ['room_number', 'hostel', 'floor', 'room_type', 'capacity', 'status']
    list_filter   = ['hostel', 'status', 'room_type']
    search_fields = ['room_number', 'hostel__name']
    list_editable = ['status']


@admin.register(Student)
class StudentAdmin(admin.ModelAdmin):
    list_display  = ['roll_number', 'name', 'department', 'year', 'phone', 'is_resident']
    list_filter   = ['department', 'year', 'is_resident']
    search_fields = ['roll_number', 'name', 'email']
    list_editable = ['is_resident']


@admin.register(RoomAllocation)
class RoomAllocationAdmin(admin.ModelAdmin):
    list_display  = ['student', 'room', 'check_in', 'check_out', 'status']
    list_filter   = ['status', 'room__hostel']
    search_fields = ['student__roll_number', 'student__name']
    raw_id_fields = ['student', 'room']
