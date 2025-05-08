import os
import traceback

# ИМПОРТЫ ИЗМЕНЕНЫ: Импортируем функции из директории handlers
from handlers.smeta_ru.processor import process_smeta_ru
from handlers.turbosmetchik.processor_1 import process_turbosmetchik_1
from handlers.turbosmetchik.processor_2 import process_turbosmetchik_2
from handlers.turbosmetchik.processor_3 import process_turbosmetchik_3

# --- Словарь для выбора функции обработки ---
PROCESSORS = {
    "Смета ру": process_smeta_ru,
    "Турбосметчик-1": process_turbosmetchik_1,
    "Турбосметчик-2": process_turbosmetchik_2,
    "Турбосметчик-3": process_turbosmetchik_3,
}

def get_available_processor_types():
    """Возвращает список ОБЩИХ типов смет для основного dropdown."""
    # Возвращаем только уникальные "основные" типы
    # Для "Турбосметчик-X" возвращаем только "Турбосметчик"
    main_types = set()
    for key in PROCESSORS.keys():
        if key.startswith("Турбосметчик-"):
            main_types.add("Турбосметчик")
        else:
            main_types.add(key)
    # Сортируем для предсказуемого порядка (опционально)
    return sorted(list(main_types))

# --- Функция-диспетчер ---
def run_processor(smeta_type, input_path):
    """
    Выбирает и запускает нужную функцию обработки на основе типа сметы.

    Args:
        smeta_type (str): Тип сметы (ключ из словаря PROCESSORS).
        input_path (str): Путь к входному файлу.

    Returns:
        tuple: (headers, data_rows) или (None, None) если произошла ошибка.
    """
    print(f"Вызов run_processor: тип={smeta_type}, файл={os.path.basename(input_path)}") # Улучшил лог
    processor_func = PROCESSORS.get(smeta_type)
    if processor_func:
        print(f"Выбран процессор: {processor_func.__name__}")
        try:
            result = processor_func(input_path)
            if isinstance(result, tuple) and len(result) == 2:
                return result
            else:
                print(f"[ОШИБКА] Обработчик '{smeta_type}' ({processor_func.__name__}) вернул некорректный результат: {result}")
                return None, None
        except Exception as e:
            print(f"[КРИТИЧЕСКАЯ ОШИБКА] Исключение при вызове '{smeta_type}' ({processor_func.__name__}) для {input_path}: {e}")
            traceback.print_exc()
            return None, None
    else:
        print(f"[ОШИБКА] Обработчик для типа '{smeta_type}' не найден в словаре PROCESSORS.")
        return None, None 