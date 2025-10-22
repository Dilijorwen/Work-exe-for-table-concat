import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def read_excel_safely(filepath):
    try:
        return pd.read_excel(filepath, engine="openpyxl", header=None)
    except Exception:
        try:
            return pd.read_excel(filepath, engine="xlrd", header=None)
        except Exception as e:
            raise RuntimeError(f"Ошибка чтения {filepath}: {e}")

def process_excels():
    filepaths = filedialog.askopenfilenames(
        title="Выберите Excel-файлы",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not filepaths:
        return

    df_list = []

    for file in filepaths:
        try:
            df = read_excel_safely(file)

            start_idx = None
            for i, val in enumerate(df.iloc[:, 0]):
                if str(val).strip().isdigit():
                    start_idx = i
                    break
            if start_idx is None:
                messagebox.showwarning("Пропуск файла", f"Не удалось найти начало таблицы:\n{os.path.basename(file)}")
                continue


            df = df.iloc[start_idx:].copy()

            itogo_index = df[df.astype(str)
                .apply(lambda x: x.str.contains("итого", case=False, na=False))
                .any(axis=1)].index
            if len(itogo_index) > 0:
                df = df.loc[:itogo_index[0] - 1]

            df = df.dropna(how='all')
            first_col = df.columns[0]
            df = df[df[first_col].astype(str).str.match(r'^\s*\d+\s*$')]
            df = df.reset_index(drop=True)

            df_list.append(df)

        except Exception as e:
            messagebox.showwarning("Ошибка файла", f"Не удалось обработать {os.path.basename(file)}:\n{e}")

    if not df_list:
        messagebox.showerror("Ошибка", "Не удалось собрать данные ни из одного файла.")
        return

    merged = pd.concat(df_list, ignore_index=True)

    date_col = None
    for col in merged.columns:
        sample = merged[col].astype(str).head(50)
        if sample.str.contains(r'\d{1,2}\.\d{1,2}\.\d{4}').any():
            date_col = col
            break
    if date_col is not None:
        merged[date_col] = pd.to_datetime(merged[date_col], errors='coerce', dayfirst=True)
        merged[date_col] = merged[date_col].dt.strftime("%d.%m.%Y")
    try:
        merged[first_col] = pd.to_numeric(merged[first_col], errors='coerce')
    except Exception:
        pass

    sort_keys = [first_col]
    if date_col is not None:
        sort_keys.append(date_col)

    merged = merged.sort_values(by=sort_keys, ascending=True, na_position='last').reset_index(drop=True)

    output_file = filedialog.asksaveasfilename(
        title="Сохранить итоговый файл как...",
        defaultextension=".csv",
        filetypes=[("CSV (для Google Sheets)", "*.csv")]
    )
    if output_file:
        try:
            merged.to_csv(output_file, index=False, header=False, encoding="utf-8-sig")
            messagebox.showinfo("Готово", f"Файл успешно сохранён в формате CSV (UTF-8):\n{output_file}")
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить файл:\n{e}")

root = tk.Tk()
root.title("Объединение Excel-файлов")
root.geometry("440x120")
root.resizable(False, False)

info = tk.Label(root, text="Выберите Excel-файлы для объединения.", justify="center")
info.pack(pady=10)

btn_select = tk.Button(root, text="Выбрать файлы и объединить", command=process_excels, font=("Arial", 12))
btn_select.pack(pady=20)

root.mainloop()