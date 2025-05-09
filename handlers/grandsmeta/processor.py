import openpyxl
import traceback
# Используем АБСОЛЮТНЫЙ импорт utils
from utils import is_likely_empty, check_merge, get_start_coord, get_item_id_nature

def process_grandsmeta_mixed(input_path): # Изменил имя функции для новой версии
    """
    ОБРАБАТЫВАЕТ один Excel файл по НОВЫМ ПРАВИЛАМ "Турбосметчик-1" (версия 3).
    - Название раздела/подраздела идет отдельной строкой.
    - Итог по разделу/подразделу идет отдельной строкой после всех позиций.
    ВОЗВРАЩАЕТ данные (заголовки и координаты) для дальнейшей обработки.
    """
    output_headers = ["№№ п/п", "Шифр расценки и коды ресурсов", "Наименование работ и затрат", "Единица измерения", "Кол-во единиц", "ВСЕГО затрат, руб."]
    start_id_col_idx = 0    # A
    item_total_cost_col_idx = 5   # F - Итоговая цена теперь в колонке F (0-based index 5)

    processed_rows_list = []
    active_items_buffer = [] # Буфер для основных позиций, ожидающих общую цену
    # pending_section_header и pending_subsection_header больше не нужны в прежнем виде
    first_section_found = False

    workbook = None
    try:
        workbook = openpyxl.load_workbook(filename=input_path, data_only=True)
        if not workbook.sheetnames:
            return None, None
        worksheet = workbook[workbook.sheetnames[0]]

        for row_num, row_cells_tuple in enumerate(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row), start=2):
            row_cells = list(row_cells_tuple)
            non_empty_info = [(i, getattr(c, 'coordinate', None), c.value) for i, c in enumerate(row_cells) if not is_likely_empty(c.value)]
            if not non_empty_info:
                continue

            cell_A = row_cells[start_id_col_idx] if len(row_cells) > start_id_col_idx else None
            cell_A_value_str = str(getattr(cell_A, 'value', '')).strip() if cell_A else ""
            cell_C = row_cells[2] if len(row_cells) > 2 else None
            cell_C_value_str = str(getattr(cell_C, 'value', '')).strip() if cell_C else ""

            row_type = None
            # Проверяем объединения для определения типа
            header_merge_AK_coord = check_merge(worksheet, row_num, 0, 10) # A(0) - K(10)
            footer_merge_CH_coord_candidate = check_merge(worksheet, row_num, 2, 7) # C(2) - H(7)

            if header_merge_AK_coord:
                if cell_A_value_str.startswith("Раздел"):
                    row_type = "section_header_name" # Новый тип
                else:
                    row_type = "subsection_header_name" # Новый тип
            elif footer_merge_CH_coord_candidate:
                if cell_C_value_str.startswith("Итого по разделу"):
                    row_type = "section_footer_total" # Новый тип
                elif cell_C_value_str.startswith("Итого по подразделу"):
                    row_type = "subsection_footer_total" # Новый тип
            elif cell_C_value_str == "Всего по позиции":
                row_type = "item_price_row"
            else:
                if cell_A and cell_A.data_type != 'f' and not is_likely_empty(cell_A.value):
                    try:
                        float(str(cell_A.value).replace(',', '.').strip())
                        row_type = "item"
                    except (ValueError, TypeError):
                        pass
            
            # Сброс буфера позиций при встрече любого хедера или футера
            if row_type in ["section_header_name", "subsection_header_name", 
                            "section_footer_total", "subsection_footer_total"] and active_items_buffer:
                if first_section_found:
                    processed_rows_list.extend(active_items_buffer)
                active_items_buffer = []

            if not first_section_found and row_type != "section_header_name":
                continue # Игнорируем все до первого заголовка раздела
            
            # --- Обработка типов строк ---
            if row_type == "section_header_name":
                first_section_found = True # Активируем флаг
                processed_rows_list.append({
                    "type": "header_name",
                    "level": "section",
                    "source_row_num": row_num,
                    "A_K_merge_coord": header_merge_AK_coord,
                    "name_text": cell_A_value_str
                })
            elif row_type == "subsection_header_name":
                processed_rows_list.append({
                    "type": "header_name",
                    "level": "subsection",
                    "source_row_num": row_num,
                    "A_K_merge_coord": header_merge_AK_coord,
                    "name_text": cell_A_value_str
                })
            elif row_type == "section_footer_total":
                cell_V_footer = row_cells[item_total_cost_col_idx] if len(row_cells) > item_total_cost_col_idx else None
                processed_rows_list.append({
                    "type": "header_footer",
                    "level": "section",
                    "source_row_num": row_num,
                    "C_H_merge_coord": footer_merge_CH_coord_candidate,
                    "footer_text_content": cell_C_value_str,
                    "total_V_coord": getattr(cell_V_footer, 'coordinate', None)
                })
            elif row_type == "subsection_footer_total":
                cell_V_footer = row_cells[item_total_cost_col_idx] if len(row_cells) > item_total_cost_col_idx else None
                processed_rows_list.append({
                    "type": "header_footer",
                    "level": "subsection",
                    "source_row_num": row_num,
                    "C_H_merge_coord": footer_merge_CH_coord_candidate,
                    "footer_text_content": cell_C_value_str,
                    "total_V_coord": getattr(cell_V_footer, 'coordinate', None)
                })
            elif row_type == "item_price_row":
                cell_V_price = row_cells[item_total_cost_col_idx] if len(row_cells) > item_total_cost_col_idx else None
                price_total_coord_for_buffer = getattr(cell_V_price, 'coordinate', None)
                for item_in_buffer in active_items_buffer:
                    item_in_buffer["col_6_coord"] = price_total_coord_for_buffer
                if first_section_found and active_items_buffer: # Проверка first_section_found здесь избыточна, т.к. item не добавился бы без него
                    processed_rows_list.extend(active_items_buffer)
                    active_items_buffer = []
            elif row_type == "item":
                # first_section_found уже должен быть True, чтобы дойти сюда
                item_id_type = get_item_id_nature(cell_A.value)
                if item_id_type == "not_a_number":
                    continue
                
                item_data = {
                    "type": "item", 
                    "source_row_num": row_num, 
                    "col_6_coord": None
                }
                # №, шифр, наименование, ед.изм, кол‑во  →  A, B, C, D, E
                input_indices_map = {1: 0, 2: 1, 3: 2, 4: 3, 5: 4} # Стало A, B, C, D, E
                for out_col_num, in_col_idx in input_indices_map.items():
                    cell_to_map = row_cells[in_col_idx] if in_col_idx < len(row_cells) else None
                    item_data[f"col_{out_col_num}_coord"] = getattr(cell_to_map, 'coordinate', None)
                
                if item_id_type == "decimal": # Материал
                    cell_V_material = row_cells[item_total_cost_col_idx] if len(row_cells) > item_total_cost_col_idx else None
                    item_data["col_6_coord"] = getattr(cell_V_material, 'coordinate', None)
                    processed_rows_list.append(item_data)
                elif item_id_type == "integer": # Основная позиция
                    # Проверяем inline-цену в колонке F (item_total_cost_col_idx = 5)
                    cell_F_inline_price = row_cells[item_total_cost_col_idx] if len(row_cells) > item_total_cost_col_idx else None
                    # Если в Грандсмете цена для основной позиции может быть объединена (например F-G),
                    # то check_merge(worksheet, row_num, item_total_cost_col_idx, item_total_cost_col_idx + 1)
                    # Если цена всегда в одной ячейке F, то merge_coord не нужен для цены.
                    # Пока оставим проверку на одну ячейку F, если нужно объединение - нужно будет уточнить.
                    merge_FG_coord = None # Пример: check_merge(worksheet, row_num, item_total_cost_col_idx, item_total_cost_col_idx + 1) # если F-G

                    if merge_FG_coord: # Если цена в объединенной ячейке F-G
                        item_data["col_6_coord_is_range"] = True
                        item_data["col_6_coord"] = merge_FG_coord 
                    elif not is_likely_empty(getattr(cell_F_inline_price, 'value', None)): # Если есть значение в F (inline)
                        item_data["col_6_coord"] = getattr(cell_F_inline_price, 'coordinate', None)
                    else:
                        active_items_buffer.append(item_data)
                        continue 

        # Добавляем оставшиеся элементы из буфера, если они есть
        if first_section_found and active_items_buffer:
            processed_rows_list.extend(active_items_buffer)

        # Сортируем все собранные строки по их исходному номеру строки
        processed_rows_list.sort(key=lambda x: x.get('source_row_num', float('inf')))
        
        all_coords_data = []
        for row_data in processed_rows_list:
            item_type = row_data.get("type")

            if item_type == "header_name":
                # ➊ координата A‑ячейки (левый‑верх диапазона A‑K)
                a_coord = get_start_coord(row_data.get("A_K_merge_coord"))  # 'A34'
                # output_headers должен быть доступен здесь, если нет, нужно его передать или определить
                # В текущем коде output_headers определен в начале функции process_grandsmeta_mixed
                header_row = [None] * len(output_headers)
                header_row[0] = a_coord                 # → колонка A
                header_row[1] = row_data.get("name_text")   # → колонка B
                # header_row[2] остаётся None по умолчанию из [None] * len(output_headers)
                all_coords_data.append(header_row)

            elif item_type == "header_footer":
                all_coords_data.append(['__FOOTER__',
                                        row_data.get("footer_text_content"),
                                        row_data.get("total_V_coord")]) # total_V_coord уже должен быть одиночной координатой

            elif item_type == "item":
                coords_row = [None] * len(output_headers) # Для item нужен этот массив
                for i_col in range(5): # Колонки 1-5 (индексы 0-4)
                    coords_row[i_col] = get_start_coord(row_data.get(f"col_{i_col+1}_coord"))
                coords_row[5] = get_start_coord(row_data.get("col_6_coord")) # Колонка 6 (ВСЕГО)
                all_coords_data.append(coords_row)

        return output_headers, all_coords_data

    except FileNotFoundError:
        return None, None
    except Exception as e:
        print(f"[КРИТИЧЕСКАЯ ОШИБКА] при обработке файла '{input_path}' (Турбосметчик-1, НОВЫЕ ПРАВИЛА v3): {e}")
        print("-" * 60); traceback.print_exc(); print("-" * 60)
        return None, None
    finally:
        if workbook:
            try: workbook.close()
            except Exception: pass