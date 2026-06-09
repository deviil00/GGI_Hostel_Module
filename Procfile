web: cd hostel && python manage.py migrate --no-input && python manage.py collectstatic --no-input && gunicorn hostel_project.wsgi --bind 0.0.0.0:$PORT
