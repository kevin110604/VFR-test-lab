from flask import session
from config import PASSWORD_STL, PASSWORD_WTL

def login(password_input):
    if password_input == PASSWORD_STL:
        session['auth_ok'] = True
        session['user_type'] = 'stl'
        return True
    elif password_input == PASSWORD_WTL:
        session['auth_ok'] = True
        session['user_type'] = 'wtl'
        return True
    return False

def logout():
    session.pop('auth_ok', None)
    session.pop('user_type', None)

def is_logged_in():
    return session.get('auth_ok', False)

def get_user_type():
    return session.get('user_type', 'stl')
