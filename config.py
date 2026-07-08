import os

def get_desktop_path():
    """Возвращает путь к рабочему столу текущего пользователя."""
    if os.name == 'nt':
        return os.path.join(os.environ.get('USERPROFILE', ''), 'Desktop')
    else:
        return os.path.join(os.path.expanduser('~'), 'Desktop')

INPUT_DIR = get_desktop_path()
OUTPUT_FILE = os.path.join('output', 'merged_result.xlsx')
ALLOWED_EXTENSIONS = {'.xlsx', '.xls'}
DEDUP_SUBSET = None
JOIN_OPTION = 'outer'