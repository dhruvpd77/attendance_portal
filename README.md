# LJIET_Attendance

A production-ready Django attendance portal with three roles: **Admin**, **Faculty**, and **Student**. Built with Bootstrap 5 and LJIET branding.

## Features

- **Admin**: Manage departments, batches, subjects, faculties, students; create schedule (who teaches what, when); set term phases (T1–T4); view daily absent report and download Excel; generate attendance sheet by batch/phase.
- **Faculty**: View today’s lectures; mark attendance by date (select absent roll numbers); download report Excel for a date/batch.
- **Student**: View today’s schedule; see attendance summary (held, attended, overall %).

## Setup

```bash
cd attendance_portal
pip install -r requirements.txt
cp .env.example .env   # optional: edit SECRET_KEY, DEBUG, ALLOWED_HOSTS
python manage.py migrate
python manage.py createsuperuser   # for Django admin and default admin role
```

## Create users with roles

- **Admin** (after creating a superuser, use Django admin to add `UserRole` with role=admin for that user, or use the command below with a new user):

  ```bash
  python manage.py create_attendance_user admin adminuser yourpassword
  ```

- **Faculty** (create Faculty in Admin or in the app, then link a user):

  ```bash
  python manage.py create_attendance_user faculty fac1 yourpassword --faculty-id=1
  ```

- **Student** (create Student in Admin or via CSV upload, then link):

  ```bash
  python manage.py create_attendance_user student stu1 yourpassword --student-id=1
  ```

## Run

```bash
python manage.py runserver
```

Open **http://127.0.0.1:8000/** and log in with a user that has a role. Superuser/staff are treated as admin and redirected to the admin dashboard. After login, all app pages (dashboard, departments, batches, etc.) are under **/portal/** (e.g. `/portal/admin-dashboard/`, `/portal/admin/departments/`).

## Quick start (data)

1. Log in as admin → set **current department** on Dashboard.
2. Add **Departments**, **Batches**, **Subjects**, **Faculties**.
3. **Upload students** (CSV with columns: roll_no, name, enrollment_no).
4. Add **Schedule** slots (faculty, subject, batch, day, time_slot).
5. Set **Term Phases** (T1–T4 start/end dates).
6. Create **faculty/student users** and link to Faculty/Student.
7. Faculty can then **Mark Attendance**; Admin can view **Daily Absent** and **Attendance Sheet** and download Excel.

## Production

```bash
export DEBUG=False
export SECRET_KEY=your-secure-secret-key
export ALLOWED_HOSTS=yourdomain.com
python manage.py collectstatic --noinput
gunicorn config.wsgi:application --bind 0.0.0.0:8000
```

## Tech

- Django 5.x, Bootstrap 5, Font Awesome
- SQLite (default); PostgreSQL via DATABASE_URL
- python-dotenv, gunicorn, whitenoise
- openpyxl for Excel export
