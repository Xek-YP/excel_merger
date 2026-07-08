import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from excel_merger import merge_excel_files
from file_loader import get_excel_files
from config import INPUT_DIR, OUTPUT_FILE, DEDUP_SUBSET, JOIN_OPTION

class ExcelMergerGUI:
    def __init__(self, root):
        self.root = root
        root.title("Excel Merger Pro")
        root.geometry("700x650")
        root.resizable(True, True)

        # Переменные
        self.input_dir = tk.StringVar(value=INPUT_DIR)
        self.output_file = tk.StringVar(value=OUTPUT_FILE)
        self.dedup_cols = tk.StringVar(value="")  # столбцы через запятую
        self.join_option = tk.StringVar(value=JOIN_OPTION)
        self.is_running = False

        # Создаём виджеты
        self.create_widgets()

        # При старте обновляем список файлов и выходной путь
        self.update_file_list()
        self._update_output_from_input()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # ---- Входная папка ----
        ttk.Label(main_frame, text="Папка с файлами:").grid(row=0, column=0, sticky="w", pady=5)
        entry_dir = ttk.Entry(main_frame, textvariable=self.input_dir, width=50)
        entry_dir.grid(row=0, column=1, sticky="ew", padx=5)
        btn_browse_dir = ttk.Button(main_frame, text="Обзор...", command=self.browse_input_dir)
        btn_browse_dir.grid(row=0, column=2, padx=5)

        # ---- Список файлов ----
        ttk.Label(main_frame, text="Найденные Excel-файлы:").grid(row=1, column=0, sticky="w", pady=5)
        self.file_listbox = tk.Listbox(main_frame, height=6)
        self.file_listbox.grid(row=2, column=0, columnspan=3, sticky="ew", pady=5)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.grid(row=2, column=3, sticky="ns")
        self.file_listbox.config(yscrollcommand=scrollbar.set)
        btn_refresh = ttk.Button(main_frame, text="Обновить список", command=self.update_file_list)
        btn_refresh.grid(row=3, column=0, columnspan=3, sticky="w", pady=2)

        # ---- Выходной файл ----
        ttk.Label(main_frame, text="Результат сохранить как:").grid(row=4, column=0, sticky="w", pady=5)
        entry_out = ttk.Entry(main_frame, textvariable=self.output_file, width=50)
        entry_out.grid(row=4, column=1, sticky="ew", padx=5)
        btn_browse_out = ttk.Button(main_frame, text="Обзор...", command=self.browse_output_file)
        btn_browse_out.grid(row=4, column=2, padx=5)

        # ---- Настройки ----
        ttk.Label(main_frame, text="Столбцы для удаления дубликатов (через запятую):").grid(row=5, column=0, sticky="w", pady=5)
        entry_dedup = ttk.Entry(main_frame, textvariable=self.dedup_cols, width=40)
        entry_dedup.grid(row=5, column=1, sticky="w", padx=5)
        ttk.Label(main_frame, text="(оставьте пустым – все столбцы)").grid(row=5, column=2, sticky="w")

        ttk.Label(main_frame, text="Тип объединения:").grid(row=6, column=0, sticky="w", pady=5)
        radio_frame = ttk.Frame(main_frame)
        radio_frame.grid(row=6, column=1, columnspan=2, sticky="w")
        ttk.Radiobutton(radio_frame, text="Все столбцы", variable=self.join_option, value="outer").pack(side="left", padx=5)
        ttk.Radiobutton(radio_frame, text="Только общие столбцы, которые есть во всех файлах", variable=self.join_option, value="inner").pack(side="left", padx=5)

        # ---- Кнопка запуска и прогресс ----
        self.btn_start = ttk.Button(main_frame, text="Запустить объединение", command=self.start_merge)
        self.btn_start.grid(row=7, column=0, columnspan=3, pady=10)

        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=400, mode="indeterminate")
        self.progress.grid(row=8, column=0, columnspan=3, pady=5, sticky="ew")

        # ---- Лог ----
        ttk.Label(main_frame, text="Лог выполнения:").grid(row=9, column=0, sticky="w", pady=5)
        self.log_text = scrolledtext.ScrolledText(main_frame, height=15, state="normal", wrap="word")
        self.log_text.grid(row=10, column=0, columnspan=3, sticky="nsew", pady=5)
        main_frame.rowconfigure(10, weight=1)
        main_frame.columnconfigure(1, weight=1)

        btn_clear_log = ttk.Button(main_frame, text="Очистить лог", command=self.clear_log)
        btn_clear_log.grid(row=11, column=0, columnspan=3, sticky="e", pady=2)

    def browse_input_dir(self):
        directory = filedialog.askdirectory(initialdir=self.input_dir.get())
        if directory:
            self.input_dir.set(directory)
            self.update_file_list()
            self._update_output_from_input()   # автоматически обновляем путь результата

    def browse_output_file(self):
        filetypes = [("Excel files", "*.xlsx"), ("All files", "*.*")]
        # начальный каталог - папка ввода, если она существует
        initial_dir = self.input_dir.get() if os.path.isdir(self.input_dir.get()) else os.path.dirname(self.output_file.get())
        filename = filedialog.asksaveasfilename(
            initialdir=initial_dir,
            defaultextension=".xlsx",
            filetypes=filetypes
        )
        if filename:
            self.output_file.set(filename)

    def update_file_list(self):
        self.file_listbox.delete(0, tk.END)
        directory = self.input_dir.get()
        if not os.path.isdir(directory):
            self.file_listbox.insert(tk.END, "Папка не существует")
            return
        files = get_excel_files(directory)
        if not files:
            self.file_listbox.insert(tk.END, "Excel-файлы не найдены")
        else:
            for f in files:
                self.file_listbox.insert(tk.END, os.path.basename(f))

    def _update_output_from_input(self):
        """Автоматически устанавливает выходной файл в ту же папку с именем merged_result.xlsx."""
        input_dir = self.input_dir.get()
        if os.path.isdir(input_dir):
            default_name = "merged_result.xlsx"
            output_path = os.path.join(input_dir, default_name)
            self.output_file.set(output_path)

    def clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")

    def log_message(self, message):
        def _append():
            self.log_text.config(state="normal")
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state="disabled")
        self.root.after(0, _append)

    def start_merge(self):
        if self.is_running:
            messagebox.showwarning("Занято", "Объединение уже выполняется. Подождите.")
            return

        input_dir = self.input_dir.get().strip()
        output_file = self.output_file.get().strip()
        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showerror("Ошибка", "Укажите корректную папку с файлами.")
            return
        if not output_file:
            messagebox.showerror("Ошибка", "Укажите путь для сохранения результата.")
            return
        files = get_excel_files(input_dir)
        if not files:
            messagebox.showerror("Ошибка", "В выбранной папке нет Excel-файлов.")
            return

        dedup_str = self.dedup_cols.get().strip()
        dedup_subset = [col.strip() for col in dedup_str.split(',') if col.strip()] if dedup_str else None
        join = self.join_option.get()

        self.is_running = True
        self.btn_start.config(state="disabled")
        self.progress.start(10)
        self.clear_log()
        self.log_message("Начинаем объединение...")

        thread = threading.Thread(target=self._run_merge, args=(input_dir, output_file, dedup_subset, join))
        thread.daemon = True
        thread.start()

    def _run_merge(self, input_dir, output_file, dedup_subset, join):
        error_msg = None
        try:
            stats = merge_excel_files(
                input_dir=input_dir,
                output_file=output_file,
                dedup_subset=dedup_subset,
                join=join,
                log_callback=self.log_message
            )
            self.log_message("\n--- ИТОГОВАЯ СТАТИСТИКА ---")
            self.log_message(f"Найдено файлов: {stats['files_found']}")
            self.log_message(f"Загружено файлов: {stats['files_loaded']}")
            self.log_message(f"Всего строк до удаления дубликатов: {stats['total_rows_before']}")
            self.log_message(f"Удалено дубликатов: {stats['duplicates_removed']}")
            self.log_message(f"Итоговое количество строк: {stats['total_rows_after']}")
            self.log_message(f"Результат сохранён: {stats['output_file']}")
            self.root.after(0, lambda: messagebox.showinfo("Готово", f"Объединение завершено!\nРезультат в {output_file}"))
        except Exception as e:
            error_msg = str(e)
            self.log_message(f"Критическая ошибка: {error_msg}")
            self.root.after(0, lambda: messagebox.showerror("Ошибка", error_msg))
        finally:
            self.root.after(0, self._finish_merge)

    def _finish_merge(self):
        self.is_running = False
        self.btn_start.config(state="normal")
        self.progress.stop()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerGUI(root)
    root.mainloop()
