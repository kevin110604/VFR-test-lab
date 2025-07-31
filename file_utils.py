import json
import os
import openpyxl

def safe_read_json(path, default=None):
    if not os.path.exists(path):
        return default if default is not None else []
    with open(path, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except Exception:
            return default if default is not None else []

def safe_write_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def safe_read_text(path, default=""):
    if not os.path.exists(path):
        return default
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

def safe_write_text(path, text):
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)

def safe_load_excel(path):
    return openpyxl.load_workbook(path)

def safe_save_excel(wb, path):
    wb.save(path)
