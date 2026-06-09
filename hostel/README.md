# Hostel Management System

Django + PostgreSQL. No extras.

## Project Structure

```
hostel/
├── apps/
│   ├── accounts/     — User model (Admin, Warden, Student roles)
│   ├── hostel/       — Hostel, Room, Student, RoomAllocation + all views
│   ├── fees/         — FeeStructure, FeeRecord
│   └── complaints/   — Complaint
├── hostel_project/
│   ├── settings.py
│   └── urls.py
├── templates/
│   ├── base/         — base.html, login.html
│   ├── admin/        — dashboard, students, rooms, fees, complaints, etc.
│   └── student/      — dashboard, raise_complaint
├── manage.py
├── requirements.txt
└── .env.example
```

## Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Create .env file
cp .env.example .env
# Edit .env — set your PostgreSQL credentials

# 3. Create the database in PostgreSQL
createdb hostel_db

# 4. Run migrations
python manage.py migrate

# 5. Create admin user
python manage.py createsuperuser

# 6. Run
python manage.py runserver
```

## URLs

| URL | Who | What |
|-----|-----|------|
| `/login/` | Everyone | Login page |
| `/dashboard/` | Admin/Warden | Admin dashboard |
| `/student/dashboard/` | Student | Student dashboard |
| `/students/` | Admin | List + search students |
| `/students/import/` | Admin | Upload Excel file |
| `/students/<id>/` | Admin | Student detail + history |
| `/students/<id>/allocate/` | Admin | Allocate room to student |
| `/rooms/` | Admin | Room list with filters |
| `/checkout/<id>/` | Admin | Check out student |
| `/fees/` | Admin | Fee records |
| `/fees/generate/` | Admin | Generate bulk fee records |
| `/complaints/` | Admin | All complaints |
| `/complaints/raise/` | Student | Raise a complaint |
| `/admin/` | Superuser | Django admin panel |

## Excel Import

Go to `/students/import/` and upload an `.xlsx` file.

**Required columns:** `Roll Number`, `Name`

**Optional columns:** `Email`, `Phone`, `Department`, `Year`, `Semester`

Column names are detected automatically (case-insensitive).
If a student with the same roll number already exists, their record is **updated**.
Room allocation is done separately by the admin after import.

## Fee Workflow

1. Admin creates a `FeeStructure` in Django Admin (`/admin/fees/feestructure/add/`)
   - Set name, academic year, semester, hostel fee, mess fee, due date
2. Admin goes to `/fees/generate/` and selects the fee structure
3. One `FeeRecord` is created for every active resident
4. Admin updates payment status at `/fees/` (amount paid, mode, reference)

## Roles

- **Admin** — full access to everything
- **Warden** — same as admin (can manage rooms, students, complaints, fees)
- **Student** — can only see their own room, fees, complaints; can raise complaints
