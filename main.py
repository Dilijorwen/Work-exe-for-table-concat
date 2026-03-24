import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd


# Какие заголовки ищем в исходных таблицах
REQUIRED_SOURCE_HEADERS = [
    "Дата",
    "Номер",
    "Поступление",
    "Списание",
    "Назначение платежа",
    "Контрагент",
    "Организация",
    "Банковский счет",
]



def normalize_text(value) -> str:
    """Нормализация текста для сравнения заголовков."""
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\n", " ")
    text = text.replace("\r", " ")
    text = re.sub(r"\s+", " ", text)
    text = text.replace("ё", "е")
    return text


def is_empty(value) -> bool:
    if value is None:
        return True
    text = str(value).strip()
    return text == "" or text.lower() == "nan"


def try_parse_amount(value):
    """Пытается привести сумму к нормальному виду строки."""
    if is_empty(value):
        return ""

    text = str(value).strip()
    text = text.replace("\xa0", " ").replace(" ", "")

    # Если число с запятой
    text = text.replace(",", ".")

    try:
        num = float(text)
        # Возвращаем красиво, но без лишних нулей
        result = f"{num:.2f}".rstrip("0").rstrip(".")
        return result.replace(".", ",")
    except Exception:
        # Если не удалось — вернуть как есть
        return str(value).strip()


def find_header_row_and_columns(df_raw: pd.DataFrame):
    """
    Ищет строку заголовков и определяет индексы нужных колонок
    по их названиям, а не по номеру.
    """
    normalized_required = {normalize_text(x): x for x in REQUIRED_SOURCE_HEADERS}

    MAX_HEADER_SCAN_ROWS = 50
    for row_idx in range(min(len(df_raw), MAX_HEADER_SCAN_ROWS)):
        row_values = df_raw.iloc[row_idx].tolist()

        found_map = {}
        for col_idx, cell_value in enumerate(row_values):
            cell_norm = normalize_text(cell_value)
            if cell_norm in normalized_required:
                original_header = normalized_required[cell_norm]
                found_map[original_header] = col_idx

        # Если нашли все нужные заголовки — это строка шапки
        if all(header in found_map for header in REQUIRED_SOURCE_HEADERS):
            return row_idx, found_map

    raise ValueError(
        "Не удалось найти строку заголовков. "
        f"Ожидались заголовки: {', '.join(REQUIRED_SOURCE_HEADERS)}"
    )

def row_to_joined_text(values) -> str:
    parts = []
    for v in values:
        if not is_empty(v):
            parts.append(str(v).strip())
    return " ".join(parts).lower().replace("ё", "е")


def is_footer_or_service_row(values) -> bool:
    text = row_to_joined_text(values)

    if not text:
        return True

    bad_fragments = [
        "итого",
        "ответственный",
        "должность",
        "подпись",
        "расшифровка подписи",
    ]

    if any(fragment in text for fragment in bad_fragments):
        return True

    compact = re.sub(r"\s+", "", text)
    if compact and all(ch in "-—_()" for ch in compact):
        return True

    return False


def extract_bank_name(bank_account_text: str) -> str:
    """
    Извлекает банк из поля 'Банковский счет'.

    Правило:
    - если одна пара кавычек -> берем её
    - если две или больше -> берем вторую
    """
    if is_empty(bank_account_text):
        return ""

    text = str(bank_account_text)

    # Ищем текст в обычных двойных кавычках
    quoted = re.findall(r'"([^"]+)"', text)
    if len(quoted) >= 2:
        return quoted[1].strip()
    if len(quoted) == 1:
        return quoted[0].strip()

    # На случай «ёлочек»
    quoted_ru = re.findall(r'«([^»]+)»', text)
    if len(quoted_ru) >= 2:
        return quoted_ru[1].strip()
    if len(quoted_ru) == 1:
        return quoted_ru[0].strip()

    return ""


def load_company_mapping(mapping_file: str) -> dict:
    ext = os.path.splitext(mapping_file)[1].lower()

    mapping = {}

    if ext in [".txt", ".csv"]:
        with open(mapping_file, "r", encoding="utf-8-sig") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue

                # Разделитель строго " - "
                if " - " in line:
                    left, right = line.split(" - ", 1)
                elif "-" in line:
                    left, right = line.split("-", 1)
                else:
                    continue

                org = left.strip()
                company = right.strip()

                if org:
                    mapping[normalize_text(org)] = company

    elif ext in [".xls", ".xlsx", ".xlsm"]:
        excel_data = pd.read_excel(mapping_file, sheet_name=None, header=None, dtype=str)

        for _, df in excel_data.items():
            for _, row in df.iterrows():
                values = [str(x).strip() for x in row.tolist() if not is_empty(x)]
                if not values:
                    continue

                # Если вся строка в одной ячейке: "Организация - Компания"
                if len(values) == 1:
                    line = values[0]
                    if " - " in line:
                        left, right = line.split(" - ", 1)
                    elif "-" in line:
                        left, right = line.split("-", 1)
                    else:
                        continue
                    org = left.strip()
                    company = right.strip()
                    if org:
                        mapping[normalize_text(org)] = company

                # Если две колонки: [Организация, Компания]
                elif len(values) >= 2:
                    org = values[0].strip()
                    company = values[1].strip()
                    if org:
                        mapping[normalize_text(org)] = company
    else:
        raise ValueError("Файл соответствий должен быть txt/csv/xls/xlsx/xlsm")

    return mapping


def read_excel_all_sheets(file_path: str):
    return pd.read_excel(file_path, sheet_name=None, header=None, dtype=str)


def process_source_file(file_path: str, company_map: dict) -> pd.DataFrame:
    sheets = read_excel_all_sheets(file_path)

    processed_frames = []

    for sheet_name, df_raw in sheets.items():
        if df_raw.empty:
            continue

        try:
            header_row_idx, col_map = find_header_row_and_columns(df_raw)
        except Exception:
            # На этом листе нужная таблица не найдена
            continue

        data_rows = df_raw.iloc[header_row_idx + 1:].copy()

        # Оставим только строки, где есть хотя бы что-то полезное
        useful_cols = [col_map[h] for h in REQUIRED_SOURCE_HEADERS if h in col_map]
        data_rows = data_rows[
            data_rows[useful_cols].apply(
                lambda row: any(not is_empty(v) for v in row.tolist()), axis=1
            )
        ].copy()

        result_rows = []

        for _, row in data_rows.iterrows():
            row_values = row.tolist()

            if is_footer_or_service_row(row_values):
                continue

            date_val = row.iloc[col_map["Дата"]]
            number_val = row.iloc[col_map["Номер"]]
            income_val = row.iloc[col_map["Поступление"]]
            expense_val = row.iloc[col_map["Списание"]]
            purpose_val = row.iloc[col_map["Назначение платежа"]]
            counterparty_val = row.iloc[col_map["Контрагент"]]
            organization_val = row.iloc[col_map["Организация"]]
            bank_account_val = row.iloc[col_map["Банковский счет"]]

            # Пропуск полностью пустых строк
            if all(
                is_empty(x)
                for x in [
                    date_val,
                    number_val,
                    income_val,
                    expense_val,
                    purpose_val,
                    counterparty_val,
                    organization_val,
                    bank_account_val,
                ]
            ):
                continue

            organization_text = "Неизвестно" if is_empty(organization_val) else str(organization_val).strip()
            bank_account_text = "Неизвестно" if is_empty(bank_account_val) else str(bank_account_val).strip()

            company = company_map.get(normalize_text(organization_text), "")
            bank_name = extract_bank_name(bank_account_text)

            result_rows.append(
                {
                    "№ п/п": "",  # заполним потом
                    "Дата": "" if is_empty(date_val) else str(date_val).strip(),
                    "Номер вх.": "" if is_empty(number_val) else str(number_val).strip(),
                    "Поступление": try_parse_amount(income_val),
                    "Списание": try_parse_amount(expense_val),
                    "Назначение платежа": "" if is_empty(purpose_val) else str(purpose_val).strip(),
                    "Китаец": "",
                    "Контрагент": "" if is_empty(counterparty_val) else str(counterparty_val).strip(),
                    "Организация": organization_text,
                    "Банковский счет": bank_account_text,
                    "Компания": company,
                    "Банк": bank_name,
                }
            )

        if result_rows:
            processed_frames.append(pd.DataFrame(result_rows))

    if not processed_frames:
        raise ValueError(
            f"В файле '{os.path.basename(file_path)}' не найдено ни одного листа "
            "с нужной таблицей."
        )

    return pd.concat(processed_frames, ignore_index=True)


def merge_files(source_files: list[str], mapping_file: str, output_csv: str):
    company_map = load_company_mapping(mapping_file)

    all_frames = []
    for file_path in source_files:
        df_part = process_source_file(file_path, company_map)
        all_frames.append(df_part)

    if not all_frames:
        raise ValueError("Нет данных для объединения.")

    final_df = pd.concat(all_frames, ignore_index=True)

    # Перенумерация
    final_df["№ п/п"] = range(1, len(final_df) + 1)

    # Гарантируем порядок столбцов
    FINAL_COLUMNS = [
        "№ п/п",
        "Дата",
        "Номер вх.",
        "Поступление",
        "Списание",
        "Назначение платежа",
        "Китаец",
        "Контрагент",
        "Организация",
        "Банковский счет",
        "Компания",
        "Банк",
    ]

    final_df = final_df[FINAL_COLUMNS]

    # Сохраняем CSV
    final_df.to_csv(output_csv, index=False, header=False, sep=";", encoding="utf-8-sig")

    return len(final_df)


class BankMergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Объединение банковских таблиц в CSV")
        self.root.geometry("820x520")

        self.source_files = []
        self.mapping_file = ""
        self.output_file = ""

        self.build_ui()

    def build_ui(self):
        title = tk.Label(
            self.root,
            text="Объединение xls/xlsx таблиц в один CSV",
            font=("Arial", 14, "bold"),
        )
        title.pack(pady=10)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame,
            text="Выбрать Excel-файлы",
            width=25,
            command=self.choose_source_files,
        ).grid(row=0, column=0, padx=5, pady=5)

        tk.Button(
            btn_frame,
            text="Выбрать файл компаний",
            width=25,
            command=self.choose_mapping_file,
        ).grid(row=0, column=1, padx=5, pady=5)

        tk.Button(
            btn_frame,
            text="Куда сохранить CSV",
            width=25,
            command=self.choose_output_file,
        ).grid(row=0, column=2, padx=5, pady=5)

        tk.Button(
            self.root,
            text="Объединить",
            width=25,
            height=2,
            bg="#4CAF50",
            fg="white",
            command=self.run_merge,
        ).pack(pady=10)

        info_frame = tk.Frame(self.root)
        info_frame.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(info_frame, text="Выбранные Excel-файлы:").pack(anchor="w")
        self.files_text = tk.Text(info_frame, height=10, wrap="word")
        self.files_text.pack(fill="x", pady=5)

        tk.Label(info_frame, text="Файл компаний:").pack(anchor="w")
        self.mapping_label = tk.Label(info_frame, text="Не выбран", fg="gray", anchor="w")
        self.mapping_label.pack(fill="x", pady=3)

        tk.Label(info_frame, text="Выходной CSV:").pack(anchor="w")
        self.output_label = tk.Label(info_frame, text="Не выбран", fg="gray", anchor="w")
        self.output_label.pack(fill="x", pady=3)


    def choose_source_files(self):
        files = filedialog.askopenfilenames(
            title="Выберите Excel-файлы",
            filetypes=[
                ("Excel files", "*.xls *.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if files:
            self.source_files = list(files)
            self.files_text.delete("1.0", "end")
            for file in self.source_files:
                self.files_text.insert("end", file + "\n")

    def choose_mapping_file(self):
        file = filedialog.askopenfilename(
            title="Выберите файл соответствий компаний",
            filetypes=[
                ("Supported files", "*.txt *.csv *.xls *.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if file:
            self.mapping_file = file
            self.mapping_label.config(text=file, fg="black")

    def choose_output_file(self):
        file = filedialog.asksaveasfilename(
            title="Сохранить CSV как",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if file:
            self.output_file = file
            self.output_label.config(text=file, fg="black")

    def validate_inputs(self):
        if not self.source_files:
            raise ValueError("Не выбраны Excel-файлы.")
        if not self.mapping_file:
            raise ValueError("Не выбран файл соответствий компаний.")
        if not self.output_file:
            raise ValueError("Не выбран путь сохранения CSV.")

    def run_merge(self):
        try:
            self.validate_inputs()

            row_count = merge_files(
                source_files=self.source_files,
                mapping_file=self.mapping_file,
                output_csv=self.output_file,
            )


            messagebox.showinfo(
                "Успех",
                f"Объединение завершено.\n\nСохранено строк: {row_count}\nФайл: {self.output_file}",
            )

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


def main():
    root = tk.Tk()
    app = BankMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()