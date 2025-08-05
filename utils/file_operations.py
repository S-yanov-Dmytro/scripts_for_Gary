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

        df_new = pd.DataFrame(results)

        if temp_save:
            temp_path = source_path.replace(".xlsx", "_temp_results.xlsx")
            if os.path.exists(temp_path):
                df_existing = pd.read_excel(temp_path)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            else:
                df_combined = df_new
            df_combined.to_excel(temp_path, index=False)
            print(f"⚠️ Временные результаты сохранены в {temp_path}")
        else:
            # Сохраняем с новым именем в ту же папку
            folder = os.path.dirname(source_path)
            filename = os.path.basename(source_path).replace(".xlsx", "_results.xlsx")
            new_path = os.path.join(folder, filename)

            if os.path.exists(new_path):
                df_existing = pd.read_excel(new_path)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            else:
                df_combined = df_new

            df_combined.to_excel(new_path, index=False)
            print(f"✅ Результаты сохранены в: {new_path}")

            # Удаляем временный файл, если был
            temp_path = source_path.replace(".xlsx", "_temp_results.xlsx")
            if os.path.exists(temp_path):
                os.remove(temp_path)

    except Exception as e:
        print(f"⛔ Ошибка при сохранении результатов: {e}")
        try:
            backup_path = os.path.join(os.path.expanduser("~"), "Desktop", "registration_backup.xlsx")
            df_new.to_excel(backup_path, index=False)
            print(f"⚠️ Результаты сохранены в резервный файл: {backup_path}")
        except Exception as backup_e:
            print(f"⛔ Критическая ошибка: не удалось сохранить даже резервную копию: {backup_e}")
