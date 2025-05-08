# utils.py
import re

def is_likely_empty(value):
    """Проверяет, является ли значение 'пустым' для целей парсинга."""
    # 0 не считается пустым
    if value == 0: return False
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
                if merged_range.min_col == start_col and merged_range.max_col == end_col:
                    return merged_range.coord  # Возвращаем координаты, например 'A5:K5'
    except Exception as e:
        # Логирование предупреждения вместо print для лучшей интеграции
        # import logging
        # logging.warning(f"Не удалось проверить merge для строки {row}, столбцы {start_col}-{end_col}. Ошибка: {e}")
        print(f"  [WARN] Не удалось проверить merge для строки {row}, столбцы {start_col}-{end_col}. Ошибка: {e}")
    return None # Если не найдено или произошла ошибка

def get_start_coord(coord_str):
    """Возвращает начальную координату из диапазона ('A1:B2' -> 'A1') или саму координату."""
    if isinstance(coord_str, str) and ':' in coord_str:
        return coord_str.split(':')[0]
    return coord_str # Возвращаем как есть, если это не диапазон или не строка

def is_zero(value):
    """Проверяет, является ли значение числовым нулем."""
    if value is None: return False
    try:
        # Заменяем запятую на точку для корректного преобразования в float
        return float(str(value).replace(',', '.').strip()) == 0.0
    except (ValueError, TypeError):
        # Если не удалось преобразовать в float, это не числовой ноль
        return False

def is_integer_like(value):
    """Проверяет, можно ли представить значение как целое число (включая '1.0')."""
    if value is None: return False
    try:
        # Заменяем запятую на точку
        float_val = float(str(value).replace(',', '.').strip())
        # Проверяем, равно ли float-значение своему целочисленному представлению
        return float_val == int(float_val)
    except (ValueError, TypeError):
        # Если не удалось преобразовать в float, это не число
        return False

# Можно добавить и другие общие утилиты сюда, если появятся