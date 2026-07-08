import os
import pandas as pd
from config import ALLOWED_EXTENSIONS

def get_excel_files(directory):
    files = []
    for f in os.listdir(directory):
        ext = os.path.splitext(f)[1].lower()
        if ext in ALLOWED_EXTENSIONS:
            files.append(os.path.join(directory, f))
    return files

def load_excel(filepath):
    try:
        engine = 'openpyxl' if filepath.endswith('.xlsx') else 'xlrd'
        df = pd.read_excel(filepath, engine=engine)
        return df
    except Exception as e:
        raise e
