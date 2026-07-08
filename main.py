# main.py
import argparse
import sys
from excel_merger import merge_excel_files
from config import INPUT_DIR, OUTPUT_FILE, DEDUP_SUBSET, JOIN_OPTION

def main():
    parser = argparse.ArgumentParser(description="Объединение Excel-файлов")
    parser.add_argument('--input', default=INPUT_DIR, help="Папка с файлами")
    parser.add_argument('--output', default=OUTPUT_FILE, help="Путь сохранения")
    parser.add_argument('--dedup-subset', nargs='+', help="Столбцы для дедупликации")
    parser.add_argument('--join', choices=['outer', 'inner'], default=JOIN_OPTION)
    args = parser.parse_args()

    dedup = args.dedup_subset if args.dedup_subset else DEDUP_SUBSET
    stats = merge_excel_files(
        input_dir=args.input,
        output_file=args.output,
        dedup_subset=dedup,
        join=args.join,
        log_callback=print
    )
    # можно вывести итоговую статистику
    print("\n--- ИТОГОВАЯ СТАТИСТИКА ---")
    for k, v in stats.items():
        print(f"{k}: {v}")

if __name__ == "__main__":
    main()
