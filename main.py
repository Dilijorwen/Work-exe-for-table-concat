import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re


def read_excel_safely(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    engine = "openpyxl" if ext == ".xlsx" else "xlrd"
    try:
        return pd.read_excel(filepath, engine=engine, header=None)
    except Exception as e:
        raise RuntimeError(f"Ошибка чтения {filepath}: {e}")


def load_company_mapping():
    txt_path = filedialog.askopenfilename(
        title="Выберите текстовый файл со списком компаний",
        filetypes=[("Text files", "*.txt")]
    )
    if not txt_path:
        return {}

    company_map = {}
    try:
        with open(txt_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "-" not in line:
                    continue
                full, short = map(str.strip, line.split("-", 1))
                company_map[full] = short
    except Exception as e:
        messagebox.showerror("Ошибка чтения файла компаний", str(e))
    return company_map


def compile_company_patterns(company_map):
    return [(re.compile(re.escape(name), re.IGNORECASE), short) for name, short in company_map.items()]


def detect_company_in_row(row, patterns):
    cell_text = str(row[8]) if len(row) > 8 else ""
    for pattern, short_name in patterns:
        if pattern.search(cell_text):
            return short_name
    return "Неизвестно"


def extract_bank_name(cell_value):
    text = str(cell_value)
    if not text or text.strip() == "nan":
        return "Неизвестно"

    matches = re.findall(r'"([^"]+)"', text)
    if matches:
        return matches[-1].strip()
    else:
        return "Неизвестно"


def clean_dataframe(df):
    df.dropna(how='all', inplace=True)
    start_idx = df[df.iloc[:, 0].astype(str).str.match(r'^\s*\d+\s*$')].index
    if start_idx.empty:
        return None
    df = df.loc[start_idx[0]:].copy()

    itogo_idx = df.apply(lambda row: row.astype(str).str.contains("итого", case=False).any(), axis=1)
    if itogo_idx.any():
        df = df.loc[:itogo_idx.idxmax() - 1]

    df = df[df.iloc[:, 0].astype(str).str.match(r'^\s*\d+\s*$')]
    df.reset_index(drop=True, inplace=True)
    return df


def process_file(filepath, patterns=None):
    try:
        df = read_excel_safely(filepath)
        df = clean_dataframe(df)
        if df is None:
            messagebox.showwarning("Пропуск файла", f"Не удалось найти начало таблицы:\n{os.path.basename(filepath)}")
            return None

        if patterns:
            df["Компания"] = df.apply(lambda row: detect_company_in_row(row, patterns), axis=1)
        else:
            df["Компания"] = "Неизвестно"

        if df.shape[1] >= 10:
            df["Банк"] = df.iloc[:, 9].apply(extract_bank_name)
        else:
            df["Банк"] = "Неизвестно"

        return df
    except Exception as e:
        messagebox.showwarning("Ошибка файла", f"Не удалось обработать {os.path.basename(filepath)}:\n{e}")
        return None


def merge_and_save(df_list):
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

    first_col = merged.columns[0]
    merged[first_col] = pd.to_numeric(merged[first_col], errors='coerce')

    sort_keys = [first_col]
    if date_col is not None:
        sort_keys.append(date_col)
    merged.sort_values(by=sort_keys, ascending=True, na_position='last', inplace=True)
    merged.reset_index(drop=True, inplace=True)

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


def process_excels():
    company_map = load_company_mapping()
    patterns = compile_company_patterns(company_map) if company_map else None

    filepaths = filedialog.askopenfilenames(
        title="Выберите Excel-файлы",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not filepaths:
        return

    df_list = [df for df in (process_file(f, patterns) for f in filepaths) if df is not None]
    merge_and_save(df_list)


root = tk.Tk()
root.title("Объединение Excel-файлов")
root.geometry("520x180")
root.resizable(False, False)

info = tk.Label(root, text="1. Выберите текстовый файл с компаниями,\n"
                           "2. затем Excel-файлы для объединения.\n",
                justify="center")
info.pack(pady=10)

btn_select = tk.Button(root, text="Выбрать файлы и объединить", command=process_excels)
btn_select.pack(pady=20)

root.mainloop()