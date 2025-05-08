# formatting.py
import openpyxl
import os
from openpyxl.styles import Alignment

# Зависимости из utils.py больше не нужны здесь

def auto_adjust_column_width(worksheet):
    """Автоматически подбирает ширину для первых 6 колонок (A-F)."""
    print(f"  Применяем автоподбор ширины для листа '{worksheet.title}'...")
    for col in worksheet.columns:
        max_length = 0
        column_letter = col[0].column_letter
        # Ограничим автоподбор первыми 6 колонками для скорости
        if column_letter not in ['A', 'B', 'C', 'D', 'E', 'F']:
            # Устанавливаем минимальную ширину для остальных колонок, если она не задана референсом
            if worksheet.column_dimensions[column_letter].width is None:
                worksheet.column_dimensions[column_letter].width = 10
            continue # Пропускаем остальные колонки

        for cell in col:
            try:
                cell_value_str = str(cell.value) if cell.value is not None else ""
                # Добавим учет длины для заголовков (первая строка)
                if cell.row == 1:
                    max_length = max(max_length, len(cell_value_str) * 1.1) # Небольшой запас для заголовка
                else:
                    max_length = max(max_length, len(cell_value_str))
            except Exception:
                pass # Игнорируем ошибки при доступе к ячейкам

        # Рассчитываем ширину с небольшим запасом, но ограничиваем минимальным и максимальным значением
        adjusted_width = min(max(max_length * 1.2 + 1, 8.43), 60) # Мин 8.43 (стандарт Excel), Макс 60
        # print(f"    Колонка {column_letter}: max_length={max_length:.2f}, adjusted_width={adjusted_width:.2f}")
        worksheet.column_dimensions[column_letter].width = adjusted_width
    print(f"  Автоподбор ширины завершен для листа '{worksheet.title}'.")


def apply_reference_widths(worksheet, widths):
    """Применяет заданные ширины к первым 6 колонкам (A-F)."""
    if widths and len(widths) >= 6:
        print(f"  Применяем референсные ширины для листа '{worksheet.title}': {widths[:6]}")
        target_columns = ['A', 'B', 'C', 'D', 'E', 'F']
        for i, col_letter in enumerate(target_columns):
            if widths[i] is not None: # Проверяем, что ширина не None
                try:
                    worksheet.column_dimensions[col_letter].width = widths[i]
                except Exception as e:
                    print(f"    [WARN] Не удалось установить ширину {widths[i]} для колонки {col_letter}: {e}")
        print(f"  Применены референсные ширины для листа '{worksheet.title}'.")
        return True # Возвращаем True, если ширины были применены
    # print(f"  Референсные ширины не найдены или некорректны для листа '{worksheet.title}'.")
    return False # Возвращаем False, если ширины не применены


def apply_formatting(worksheet):
    """Применяет форматирование (центр, перенос текста) к первым 6 колонкам."""
    print(f"  Применяем форматирование (центр, перенос) для листа '{worksheet.title}'...")
    center_wrap_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    try:
        # Применяем только к первым 6 колонкам (A-F) для производительности
        for row in worksheet.iter_rows(max_col=6):
            for cell in row:
                # Проверяем, что ячейка не пустая, чтобы не форматировать лишнее
                # (хотя форматирование пустых ячеек обычно безопасно)
                # if cell.value is not None:
                cell.alignment = center_wrap_alignment
        print(f"  Форматирование применено для листа '{worksheet.title}'.")
    except Exception as e:
        print(f"  [WARN] Не удалось применить форматирование к листу '{worksheet.title}': {e}")


def read_reference_widths(reference_file_path):
    """
    Читает ширины первых 6 столбцов (A-F) из референсного файла.

    Args:
        reference_file_path (str): Путь к референсному Excel файлу.

    Returns:
        list or None: Список ширин [A, B, C, D, E, F] или None в случае ошибки или отсутствия файла.
    """
    if not os.path.exists(reference_file_path):
        print(f"  [INFO] Референсный файл не найден: {reference_file_path}. Будет использован автоподбор.")
        return None

    print(f"  Чтение референсных ширин из: {os.path.basename(reference_file_path)}...")
    wb_ref = None # Инициализируем переменную заранее
    try:
        # Используем импортированный openpyxl
        wb_ref = openpyxl.load_workbook(filename=reference_file_path, data_only=True, keep_vba=False)
        # Предполагаем, что нужный лист - активный
        ws_ref = wb_ref.active
        widths = []
        target_columns = ['A', 'B', 'C', 'D', 'E', 'F']
        for col_letter in target_columns:
            # Получаем ширину из словаря dimensions
            width = ws_ref.column_dimensions[col_letter].width if col_letter in ws_ref.column_dimensions else None
            # Используем стандартную ширину Excel (8.43), если ширина не задана явно (width is None)
            # В openpyxl width=None означает стандартную ширину, но для явной установки лучше использовать число.
            widths.append(width if width is not None else 8.43)

        print(f"  Референсные ширины прочитаны: {widths}")
        return widths
    except Exception as e:
        print(f"  [WARN] Ошибка чтения референсного файла '{os.path.basename(reference_file_path)}': {e}")
        return None # Возвращаем None при любой ошибке чтения
    finally:
        # Гарантированно закрываем файл, если он был открыт
        if wb_ref:
            try:
                wb_ref.close()
            except Exception as close_e:
                print(f"  [WARN] Не удалось закрыть референсный файл: {close_e}")