"""
accounts/views.py
Only handles the post-login redirect based on role.
All auth (login/logout) is done by Django's built-in views.
"""
# Nothing custom needed — Django's LoginView handles login.
# The dashboard view in hostel/views.py handles role-based redirect.
