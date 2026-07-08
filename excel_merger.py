import os
import pandas as pd
from file_loader import get_excel_files, load_excel
from data_combiner import combine_dataframes
from deduplicator import remove_duplicates

def merge_excel_files(input_dir, output_file, dedup_subset=None, join='outer', log_callback=print):
    """
    Основная функция объединения Excel-файлов.
    log_callback – функция для вывода сообщений (по умолчанию print).
    Возвращает словарь со статистикой.
    """
    stats = {
        'files_found': 0,
        'files_loaded': 0,
        'total_rows_before': 0,
        'total_rows_after': 0,
        'duplicates_removed': 0,
        'output_file': output_file
    }

    # 1. Получаем список файлов
    file_paths = get_excel_files(input_dir)
    stats['files_found'] = len(file_paths)
    log_callback(f"Найдено Excel-файлов: {len(file_paths)}")

    if not file_paths:
        log_callback("Ошибка: нет файлов для обработки.")
        return stats

    # 2. Загружаем все файлы
    dataframes = []
    for path in file_paths:
        try:
            df = load_excel(path)
            dataframes.append(df)
            stats['files_loaded'] += 1
            log_callback(f"  Загружен: {os.path.basename(path)} ({len(df)} строк)")
        except Exception as e:
            log_callback(f"  Ошибка загрузки {os.path.basename(path)}: {e}")

    if not dataframes:
        log_callback("Не загружено ни одного файла. Завершение.")
        return stats

    # 3. Объединяем
    try:
        combined = combine_dataframes(dataframes, join=join)
        total_before = sum(len(df) for df in dataframes)
        stats['total_rows_before'] = total_before
        log_callback(f"Объединено: {len(combined)} строк (всего до удаления дубликатов: {total_before})")
    except Exception as e:
        log_callback(f"Ошибка при объединении: {e}")
        return stats

    # 4. Удаляем дубликаты (если нужно)
    if dedup_subset is not None and len(dedup_subset) == 0:
        dedup_subset = None  # пустой список означает все столбцы
    cleaned, removed = remove_duplicates(combined, subset=dedup_subset)
    stats['duplicates_removed'] = removed
    stats['total_rows_after'] = len(cleaned)
    log_callback(f"Удалено дубликатов: {removed}")
    log_callback(f"Итоговое количество строк: {len(cleaned)}")

    # 5. Сохраняем результат
    try:
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        cleaned.to_excel(output_file, index=False, engine='openpyxl')
        log_callback(f"Результат сохранён в {output_file}")
    except Exception as e:
        log_callback(f"Ошибка сохранения: {e}")

    return stats
