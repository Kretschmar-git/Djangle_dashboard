#!/usr/bin/env bash
set -o errexit

python manage.py migrate --no-input
python manage.py create_demo_users
exec gunicorn dashboard_project.wsgi:application
