import glob
import os
from collections import Counter
from typing import Dict, List, Optional, Set, Union

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.worksheet import Worksheet
from pandas.tseries.offsets import DateOffset


def read_source_excel(file_path: str, sheet_name: str) -> Optional[pd.DataFrame]:
    """
    Читает данные из исходного Excel файла.

    Args:
        file_path (str): Путь к Excel файлу
        sheet_name (str): Название листа для чтения

    Returns:
        Optional[pd.DataFrame]: DataFrame с данными или None в случае ошибки
    """
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"Ошибка при чтении файла: {e}")
        return None


def map_support_level(warranty: str) -> str:
    """
    Преобразует значение гарантии в соответствующий уровень поддержки.

    Args:
        warranty (str): Исходное значение гарантии

    Returns:
        str: Уровень поддержки
    """
    if pd.isna(warranty):
        return "Не найдено"
    w = warranty.lower()
    if "notfound" in w:
        return "Не найдено"
    if "гарантия" in w:
        return "Гарантия"
    if w.startswith("base"):
        if "невозврат" in w:
            return "Базовый+невозврат"
        return "Базовый"
    if w.startswith("extended"):
        if "невозврат" in w:
            return "Расширенный+невозврат"
        return "Расширенный"
    if w.startswith("premium"):
        if "невозврат" in w:
            return "Премиум+невозврат"
        return "Премиум"
    return "Не найдено"


def format_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Форматирует данные из DataFrame под конечный формат.

    Args:
        df (pd.DataFrame): Исходный DataFrame с данными

    Returns:
        pd.DataFrame: Отформатированный DataFrame с нужными столбцами
    """
    # Срок: только цифра
    df["Срок"] = df["Warranty"].str.extract(r"(\d+)Y")[0]

    # SN OY: всегда строка, убираем .0 на конце
    def clean_sn(sn: Union[str, float, int]) -> str:
        """
        Очищает серийный номер от лишних символов.

        Args:
            sn (Union[str, float, int]): Исходный серийный номер

        Returns:
            str: Очищенный серийный номер
        """
        s = str(sn)
        if s.endswith(".0"):
            s = s[:-2]
        return s

    df["SN"] = df["SN"].apply(clean_sn)
    # Даты: формат YYYY-MM-DD, некорректные значения превращаем в пустую строку
    start_dates = pd.to_datetime(df["Начало гарантии"], errors="coerce")
    df["Дата начала"] = start_dates.dt.strftime("%Y-%m-%d").fillna("")

    # Дата окончания: дата начала + срок (в годах)
    def calc_end_date(row: pd.Series) -> str:
        """
        Рассчитывает дату окончания гарантии.

        Args:
            row (pd.Series): Строка данных с датой начала и сроком

        Returns:
            str: Дата окончания в формате YYYY-MM-DD или пустая строка в случае ошибки
        """
        try:
            start = pd.to_datetime(row["Дата начала"], errors="coerce")
            years = int(row["Срок"])
            if pd.isna(start) or pd.isna(years):
                return ""
            end = start + DateOffset(years=years)
            return end.strftime("%Y-%m-%d")
        except Exception:
            return ""

    df["Дата окончания"] = df.apply(calc_end_date, axis=1)
    # Форматируем уровень поддержки
    df["Уровень поддержки"] = df["Warranty"].apply(map_support_level)
    # Формируем итоговый датафрейм с нужными столбцами и переименованием
    result = pd.DataFrame(
        {
            "Наименование": df["Номенклатура"],
            "SN OY": df["SN"],
            "Уровень поддержки": df["Уровень поддержки"],
            "Срок": df["Срок"],
            "Дата начала": df["Дата начала"],
            "Дата окончания": df["Дата окончания"],
        }
    )
    return result


def check_duplicates(new_data: pd.DataFrame, target_file: str) -> pd.DataFrame:
    """
    Проверяет наличие дубликатов в целевом файле по SN OY.

    Args:
        new_data (pd.DataFrame): Новые данные для добавления
        target_file (str): Путь к целевому файлу

    Returns:
        pd.DataFrame: Объединенные данные без дубликатов
    """
    try:
        existing_data = pd.read_excel(target_file)
        combined_data = pd.concat([existing_data, new_data])
        # Оставляем только уникальные SN OY, оставляя первую встреченную строку
        combined_data = combined_data.drop_duplicates(subset=["SN OY"])
        return combined_data
    except FileNotFoundError:
        return new_data


def save_to_excel(df: pd.DataFrame, target_file: str) -> None:
    """
    Сохраняет данные в целевой Excel файл.

    Args:
        df (pd.DataFrame): Данные для сохранения
        target_file (str): Путь к целевому файлу
    """
    try:
        df.to_excel(target_file, index=False)
        print(f"Данные успешно сохранены в файл {target_file}")
    except Exception as e:
        print(f"Ошибка при сохранении файла: {e}")


def colorize_excel(target_file: str) -> None:
    """
    Применяет цветовое форматирование к Excel файлу.

    Args:
        target_file (str): Путь к целевому файлу
    """
    # Цвета для уровня поддержки
    support_colors: Dict[str, str] = {
        "Базовый": "ADD8E6",  # светло-синий
        "Базовый+невозврат": "87CEEB",  # синий
        "Гарантия": "90EE90",  # зелёный
        "Премиум": "FFFF99",  # жёлтый
        "Премиум+невозврат": "FFA500",  # оранжевый
        "Расширенный": "DDA0DD",  # фиолетовый
        "Расширенный+невозврат": "FF6347",  # красный
        "Не найдено": "D3D3D3",  # серый
    }
    # Цвета для срока
    term_colors: Dict[str, str] = {
        "1": "CCFFCC",  # светло-зелёный
        "3": "FFFFCC",  # светло-жёлтый
        "5": "FFE4B5",  # светло-оранжевый
    }
    # Цвет для дубликатов SN OY
    duplicate_sn_color: str = "FFC7CE"  # светло-красный

    wb = load_workbook(target_file)
    ws: Worksheet = wb.active

    # Найдём индексы нужных столбцов по заголовкам
    header: Dict[str, int] = {
        cell.value: idx
        for idx, cell in enumerate(next(ws.iter_rows(min_row=1, max_row=1)), 1)
    }
    col_support = header.get("Уровень поддержки")
    col_term = header.get("Срок")
    col_sn = header.get("SN OY")

    # Соберём все значения SN OY для поиска дубликатов
    sn_values: List[str] = [
        row[col_sn - 1].value for row in ws.iter_rows(min_row=2) if col_sn
    ]
    sn_counts = Counter(sn_values)
    duplicate_sns: Set[str] = {sn for sn, count in sn_counts.items() if count > 1}

    # Применяем цвета
    for row in ws.iter_rows(min_row=2):
        # Уровень поддержки
        if col_support:
            val = row[col_support - 1].value
            color = support_colors.get(val, None)
            if color:
                row[col_support - 1].fill = PatternFill(
                    start_color=color, end_color=color, fill_type="solid"
                )
        # Срок
        if col_term:
            val = str(row[col_term - 1].value)
            color = term_colors.get(val, None)
            if color:
                row[col_term - 1].fill = PatternFill(
                    start_color=color, end_color=color, fill_type="solid"
                )
        # Дубликаты SN OY
        if col_sn:
            val = row[col_sn - 1].value
            if val in duplicate_sns:
                row[col_sn - 1].fill = PatternFill(
                    start_color=duplicate_sn_color,
                    end_color=duplicate_sn_color,
                    fill_type="solid",
                )
    wb.save(target_file)


def main() -> None:
    """
    Основная функция программы.
    Обрабатывает Excel файлы в папке input и создает отформатированный файл в output.
    """
    # Пути к файлам
    # Скрипт обрабатывает все Excel-файлы в папке input
    input_dir: str = "input"
    output_dir: str = "output"
    sheet_name: str = "Гарантия"
    target_file: str = f"{output_dir}/target.xlsx"

    # Создаем папки input и output если они не существуют
    for directory in [input_dir, output_dir]:
        if not os.path.exists(directory):
            try:
                os.makedirs(directory)
                print(f"Создана папка {directory}")
            except Exception as e:
                print(f"Ошибка при создании папки {directory}: {e}")
                return

    # Удаляем существующий target.xlsx если он есть
    if os.path.exists(target_file):
        try:
            os.remove(target_file)
            print(f"Удален существующий файл {target_file}")
        except Exception as e:
            print(f"Ошибка при удалении файла {target_file}: {e}")
            return

    # Собираем все Excel-файлы в input
    input_files: List[str] = glob.glob(f"{input_dir}/*.xls*")
    if not input_files:
        print(f"Нет Excel-файлов в папке {input_dir}!")
        return

    # Создаём новый файл с нужной шапкой
    columns: List[str] = [
        "Наименование",
        "SN OY",
        "Уровень поддержки",
        "Срок",
        "Дата начала",
        "Дата окончания",
    ]
    pd.DataFrame(columns=columns).to_excel(target_file, index=False)
    print(f"Создан новый файл {target_file} с нужной структурой.")

    # Читаем и объединяем данные из всех файлов
    all_data: List[pd.DataFrame] = []
    for file in input_files:
        print(f"Чтение файла: {file}")
        source_data = read_source_excel(file, sheet_name)
        if source_data is None:
            print(f"Пропущен файл {file} (нет листа '{sheet_name}' или ошибка чтения)")
            continue
        formatted_data = format_data(source_data)
        all_data.append(formatted_data)

    if not all_data:
        print("Нет данных для обработки!")
        return

    # Объединяем все данные в один DataFrame
    combined_data: pd.DataFrame = pd.concat(all_data, ignore_index=True)

    # Проверяем дубликаты и получаем финальный набор данных
    final_data: pd.DataFrame = check_duplicates(combined_data, target_file)

    # Сохраняем результат
    save_to_excel(final_data, target_file)

    # Цветовое форматирование
    colorize_excel(target_file)


if __name__ == "__main__":
    main()
