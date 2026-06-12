from django.core.management.base import BaseCommand
from apps.hostel.models import RoomAmenity

STANDARD_AMENITIES = [
    ('Bed',       'bi-person-bed',          'Single/double bed frame'),
    ('Chair',     'bi-person-seat',         'Sitting chair'),
    ('Table',     'bi-table',               'Study/work table'),
    ('Fan',       'bi-fan',                 'Ceiling or table fan'),
    ('Tube Light','bi-lightbulb',           'Fluorescent tube light'),
    ('Mattress',  'bi-stack',               'Bed mattress'),
    ('Curtains',  'bi-window-sidebar',      'Window curtains'),
    ('AC',        'bi-wind',                'Air conditioner unit'),
    ('Cooler',    'bi-thermometer-snow',    'Air cooler'),
    ('Geyser',    'bi-thermometer-sun',     'Water heater / geyser'),
]


class Command(BaseCommand):
    help = 'Seed standard room amenity types (Bed, Chair, Table, Fan, etc.)'

    def handle(self, *args, **options):
        created = 0
        for name, icon, desc in STANDARD_AMENITIES:
            obj, was_created = RoomAmenity.objects.get_or_create(
                name=name,
                defaults={'icon': icon, 'description': desc, 'created_by': None},
            )
            if was_created:
                created += 1
                self.stdout.write(f'  Created: {name}')
            else:
                self.stdout.write(f'  Exists:  {name}')
        self.stdout.write(self.style.SUCCESS(f'\nDone — {created} new amenity type(s) seeded.'))
