from django import forms
from apps.hostel.models import Hostel, Room, Student, RoomAllocation


class HostelForm(forms.ModelForm):
    class Meta:
        model  = Hostel
        fields = ['name', 'type', 'address', 'warden']


class RoomForm(forms.ModelForm):
    class Meta:
        model  = Room
        fields = ['hostel', 'room_number', 'floor', 'room_type', 'capacity']


class AllocateRoomForm(forms.Form):
    student   = forms.ModelChoiceField(
        queryset=Student.objects.none(),   # overridden in __init__
        empty_label='Select student...',
        widget=forms.Select(attrs={'class': 'form-select'})
    )
    room      = forms.ModelChoiceField(
        queryset=Room.objects.none(),      # overridden in __init__
        empty_label='Select a room...',
        widget=forms.Select(attrs={'class': 'form-select'})
    )
    check_in  = forms.DateField(
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'})
    )
    notes     = forms.CharField(
        required=False,
        widget=forms.Textarea(attrs={'rows': 2, 'class': 'form-control','placeholder': 'Any remarks...'})
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        from apps.hostel.models import RoomAllocation
        # Show ALL students who don't already have an active allocation
        allocated_ids = RoomAllocation.objects.filter(
            status='active'
        ).values_list('student_id', flat=True)
        self.fields['student'].queryset = Student.objects.exclude(
            pk__in=allocated_ids
        ).order_by('name')
        # Only vacant rooms
        self.fields['room'].queryset = Room.objects.filter(
            status='vacant'
        ).select_related('hostel').order_by('hostel__name', 'floor', 'room_number')


class CheckoutForm(forms.Form):
    check_out = forms.DateField(
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'})
    )
    notes     = forms.CharField(
        required=False, label='Reason / Notes',
        widget=forms.Textarea(attrs={'rows': 2, 'class': 'form-control'})
    )


class StudentImportForm(forms.Form):
    file = forms.FileField(
        label='Excel File (.xlsx)',
        widget=forms.ClearableFileInput(attrs={'accept': '.xlsx,.xls'})
    )

    def clean_file(self):
        f = self.cleaned_data['file']
        if not f.name.endswith(('.xlsx', '.xls')):
            raise forms.ValidationError('Only .xlsx or .xls files are accepted.')
        if f.size > 5 * 1024 * 1024:  # 5 MB limit
            raise forms.ValidationError('File size must be under 5 MB.')
        return f
