# utils.py
import re

def is_likely_empty(value):
    """Проверяет, является ли значение 'пустым' для целей парсинга."""
    # 0 не считается пустым
    if value == 0 and (isinstance(value, int) or isinstance(value, float)): # Добавил проверку типа для 0
        return False
    return value is None or str(value).strip() == ""

def check_merge(worksheet, row, start_col_idx, end_col_idx):
    """
    Проверяет, попадает ли ячейка в указанной строке (row)
    и диапазоне столбцов (start_col_idx - end_col_idx) в объединенную ячейку.
    Возвращает координату объединенной ячейки (e.g., 'A1:K1') или None.
    """
    start_col = start_col_idx + 1  # Индексы openpyxl начинаются с 1
    end_col = end_col_idx + 1
    try:
        # Проходим по всем диапазонам объединенных ячеек на листе
        for merged_range in worksheet.merged_cells.ranges:
            # Проверяем, входит ли наша строка в диапазон строк объединенной ячейки
            if merged_range.min_row <= row <= merged_range.max_row:
                # Проверяем, совпадают ли начальный и конечный столбцы
                # ИЛИ если объединенная ячейка охватывает наш диапазон (для заголовков, например)
                if (merged_range.min_col == start_col and merged_range.max_col == end_col) or \
                   (merged_range.min_col <= start_col and merged_range.max_col >= end_col): # Добавлено условие охвата
                    return merged_range.coord  # Возвращаем координаты, например 'A5:K5'
    except AttributeError: # merged_cells может отсутствовать у некоторых объектов worksheet (редко)
        # import logging
        # logging.warning(f"Атрибут 'merged_cells' отсутствует у worksheet. Не удалось проверить merge для строки {row}.")
        print(f"  [WARN] Атрибут 'merged_cells' отсутствует у worksheet. Не удалось проверить merge для строки {row}.")
    except Exception as e:
        # import logging
        # logging.warning(f"Не удалось проверить merge для строки {row}, столбцы {start_col}-{end_col}. Ошибка: {e}")
        print(f"  [WARN] Не удалось проверить merge для строки {row}, столбцы {start_col}-{end_col}. Ошибка: {e}")
    return None

def get_start_coord(coord_str):
    """Возвращает начальную координату из диапазона ('A1:B2' -> 'A1') или саму координату."""
    if isinstance(coord_str, str) and ':' in coord_str:
        return coord_str.split(':')[0]
    return coord_str

def is_zero(value):
    """Проверяет, является ли значение числовым нулем."""
    if is_likely_empty(value): # Если значение пустое, оно не ноль
        return False
    try:
        return float(str(value).replace(',', '.').strip()) == 0.0
    except (ValueError, TypeError):
        return False

def is_integer_like(value):
    """Проверяет, можно ли представить значение как целое число (включая '1.0')."""
    if is_likely_empty(value): # Пустое значение не является числом
        return False
    try:
        val_str = str(value).replace(',', '.').strip()
        float_val = float(val_str)
        # Дополнительная проверка, чтобы "1." или "1, " не считались целыми, если они не парсятся как "1.0"
        if val_str.endswith('.') or val_str.endswith(','):
            if val_str[:-1].isdigit(): # "1." -> "1"
                return float(val_str[:-1]) == int(float(val_str[:-1]))
            return False # "1.abc" или что-то такое
        return float_val == int(float_val)
    except (ValueError, TypeError):
        return False

# НОВАЯ ФУНКЦИЯ
def get_item_id_nature(cell_value_from_excel):
    """
    Определяет, является ли значение в ячейке целым числом (или его эквивалентом типа 1.0),
    дробным числом, или не числом.
    Возвращает: "integer", "decimal", "not_a_number".
    """
    if is_likely_empty(cell_value_from_excel):
        return "not_a_number"

    val_str = str(cell_value_from_excel).strip()
    
    # Убираем возможные пробелы после запятой/точки перед цифрами, если они есть,
    # например "1. 2" -> "1.2" или "1, 2" -> "1,2"
    val_str = re.sub(r'(?<=[.,])\s*(?=\d)', '', val_str)
    
    # Заменяем запятую на точку для float()
    val_str_for_float = val_str.replace(',', '.')

    try:
        num = float(val_str_for_float)
        # Проверяем, есть ли не-нулевая дробная часть
        # Используем небольшую погрешность для сравнения float
        if abs(num - int(num)) < 1e-9: # 1e-9 это маленькое число (эпсилон)
            return "integer"  # Например, 1.0, 2.0000000001, 3
        else:
            # Дополнительная проверка: если исходная строка не содержала точку/запятую,
            # но float() дает дробное (очень редкий случай, но для безопасности)
            # или если она содержала явный десятичный разделитель.
            if '.' in val_str_for_float or ',' in val_str: # Убедимся, что это было задумано как дробное
                 return "decimal" # Например, 1.1, 2.3
            # Если это число типа 1000, которое из-за каких-то региональных настроек Excel
            # могло быть записано как "1 000" и после strip() стать "1 000",
            # но при этом float("1 000") -> ValueError.
            # Однако, наша предыдущая обработка с re.sub и .replace должна была это учесть.
            # Если же это было, например, просто текстовое "1 2" (два числа через пробел),
            # float("1 2") даст ошибку, и мы перейдем в except.
            # Эта ветка маловероятна, если первичная обработка val_str сработала.
            return "not_a_number" # Если float() создал дробь, но нет явного разделителя в строке

    except (ValueError, TypeError):
        # Если не удалось преобразовать в float, это точно не число
        # или это строка типа "1-2-3" (часть шифра), которую float не поймет.
        # Попробуем более мягкую проверку на "похожесть" на числовой префикс,
        # который может быть частью шифра, но нам нужно отличить его от просто текста.
        
        # Проверка на простой числовой префикс (целый или дробный)
        # Например, "1.1 Абвгд" -> "decimal" (для ID), "1 Абвгд" -> "integer" (для ID)
        # "Раздел 1" -> "not_a_number" (т.к. "Раздел" не парсится)
        
        match_decimal = re.match(r'^(\d+([.,]\d+)?)\b', val_str) # \b - граница слова, чтобы не было "1.1абв"
        if match_decimal:
            num_part_str = match_decimal.group(1).replace(',', '.')
            try:
                num = float(num_part_str)
                if abs(num - int(num)) < 1e-9:
                    return "integer"
                else:
                    return "decimal"
            except (ValueError, TypeError):
                return "not_a_number" # Очень странный случай, если regex нашел, а float нет

        return "not_a_number"


# Можно добавить и другие общие утилиты сюда, если появятся