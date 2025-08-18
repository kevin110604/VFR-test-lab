from flask import session
from config import PASSWORD_STL, PASSWORD_WTL, PASSWORD_VFR3, MASTER_PASSWORD

def login(password_input: str) -> bool:
    # Chuẩn hoá
    pw = (password_input or "").strip()
    if pw == MASTER_PASSWORD:
        session['auth_ok'] = True
        session['user_type'] = 'superadmin'
        return True
    if pw == PASSWORD_STL:
        session['auth_ok'] = True
        session['user_type'] = 'stl'
        return True
    if pw == PASSWORD_WTL:
        session['auth_ok'] = True
        session['user_type'] = 'wtl'
        return True
    if pw == PASSWORD_VFR3:
        session['auth_ok'] = True
        session['user_type'] = 'vfr3'
        return True
    return False

def get_user_type() -> str:
    # Mặc định là 'wtl' (ít quyền), nếu chưa login
    return session.get('user_type', 'wtl')