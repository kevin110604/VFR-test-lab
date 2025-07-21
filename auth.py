from flask import session
from config import PASSWORD

def login(password_input):
    if password_input == PASSWORD:
        session['auth_ok'] = True
        return True
    return False

def logout():
    session.pop('auth_ok', None)

def is_logged_in():
    return session.get('auth_ok', False)
