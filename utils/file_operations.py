import os
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

        column_names = [
            'Title', 'First Name', 'Last Name', 'Age', 'Address Line 1',
            'Town/City', '', 'Locality', 'County', 'Postcode',
            'Phone', 'Email', 'DOB', 'Password', '...', 'Status', 'Message'
        ]

        df_new = pd.DataFrame(results, columns=column_names[:len(results[0])])

        unique_keys = ['Title', 'First Name', 'Last Name', 'Email', 'DOB']

        if temp_save:
            temp_path = source_path.replace(".xlsx", "_temp_results.xlsx")
            if os.path.exists(temp_path):
                df_existing = pd.read_excel(temp_path)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                df_combined.drop_duplicates(subset=unique_keys, keep='last', inplace=True)
            else:
                df_combined = df_new

            df_combined.to_excel(temp_path, index=False)
            print(f"⚠️ Temporary results saved to {temp_path}")

        else:
            folder = os.path.dirname(source_path)
            original_filename = os.path.basename(source_path)
            new_filename = f"result_{original_filename}"
            new_path = os.path.join(folder, new_filename)

            if os.path.exists(new_path):
                df_existing = pd.read_excel(new_path)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                df_combined.drop_duplicates(subset=unique_keys, keep='last', inplace=True)
            else:
                df_combined = df_new

            df_combined.to_excel(new_path, index=False)
            print(f"✅ Results saved to: {new_path}")

            temp_path = source_path.replace(".xlsx", "_temp_results.xlsx")
            if os.path.exists(temp_path):
                os.remove(temp_path)

    except Exception as e:
        print(f"⛔ Error while saving results: {e}")
        try:
            backup_path = os.path.join(os.path.expanduser("~"), "Desktop", "registration_backup.xlsx")
            df_new.to_excel(backup_path, index=False)
            print(f"⚠️ Results saved to backup file: {backup_path}")
        except Exception as backup_e:
            print(f"⛔ Critical error: failed to even save the backup copy: {backup_e}")
