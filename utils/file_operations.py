import os
from datetime import datetime

import pandas as pd
from tkinter import filedialog, messagebox

def load_xls_file():
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if not path:
        messagebox.showwarning("Warning", "File not selected!")
        return None, None
    try:
        df = pd.read_excel(path, header=None, dtype=str)
        if df.empty:
            raise ValueError("The file is empty!")
        return df.drop_duplicates().fillna("").values.tolist(), path
    except Exception as e:
        messagebox.showerror("Error", f"Could not load the file: {e}")
        return None, None


def save_results_to_excel(results, source_path, temp_save=False):
    try:
        if not results:
            return

        df = pd.DataFrame(results)
        if temp_save:
            temp_path = source_path.replace(".xlsx", "_temp_results.xlsx")
            df.to_excel(temp_path, index=False, header=False)
            print(f"⚠️ Временные результаты сохранены в {temp_path}")
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_path = source_path.replace(".xlsx", f"_results_{timestamp}.xlsx")

            counter = 1
            while os.path.exists(new_path):
                new_path = source_path.replace(".xlsx", f"_results_{timestamp}_{counter}.xlsx")
                counter += 1

            df.to_excel(new_path, index=False, header=False)
            print(f"✅ Результаты сохранены в {new_path}")

            temp_path = source_path.replace(".xlsx", "_temp_results.xlsx")
            if os.path.exists(temp_path):
                os.remove(temp_path)
    except Exception as e:
        print(f"⛔ Ошибка при сохранении результатов: {e}")
        try:
            backup_path = os.path.join(os.path.expanduser("~"), "Desktop", "registration_backup.xlsx")
            df.to_excel(backup_path, index=False, header=False)
            print(f"⚠️ Результаты сохранены в резервный файл: {backup_path}")
        except Exception as backup_e:
            print(f"⛔ Критическая ошибка: не удалось сохранить даже резервную копию: {backup_e}")