# handlers/smeta_ru/handler.py
import openpyxl
import traceback
# Используем АБСОЛЮТНЫЙ импорт для доступа к utils.py из корневой папки
from utils import is_likely_empty, check_merge, get_start_coord, is_zero # Предполагается, что эти utils идентичны или совместимы

# Вспомогательная функция для определения природы ID (целый/дробный)
def get_item_id_nature(cell_value_from_excel):
    val_str = str(cell_value_from_excel).strip()
    try:
        num = float(val_str.replace(',', '.'))
        if int(num) == num:
            return "integer"
        else:
            return "decimal"
    except (ValueError, TypeError):
        return "not_a_number"

def process_smeta_ru(input_path):
    """
    ОБРАБАТЫВАЕТ один Excel файл по логике "Смета ру".
    НЕ СОХРАНЯЕТ ФАЙЛ, а ВОЗВРАЩАЕТ данные для дальнейшей обработки.
    """
    # print(f"\n--- Обработка файла (Смета ру): {os.path.basename(input_path)} ---")

    output_headers = ["№№ п/п", "Шифр расценки и коды ресурсов", "Наименование работ и затрат", "Единица измерения", "Кол-во единиц", "ВСЕГО затрат, руб."]
    price_total_col_idx = 8 # I (Итоговая цена для ОСНОВНЫХ позиций, цена для ФУТЕРОВ разделов/подразделов)
    item_individual_price_col_idx = 9 # J (Индивидуальная цена ресурса, цена для МАТЕРИАЛОВ)
    price_ztr_col_idx = 10  # K (Стоимость единицы)
    start_id_col_idx = 0    # A (Номер п/п или начало шифра)

    price_row_must_be_non_empty_indices = {price_total_col_idx, price_ztr_col_idx}
    price_row_must_be_empty_indices = set(range(8))

    processed_rows_list = []
    active_items_buffer = []
    pending_section_header = None
    pending_subsection_header = None
    first_section_found = False
    skipped_items_zero_j_count = 0

    workbook = None

    try:
        workbook = openpyxl.load_workbook(filename=input_path, data_only=True)
        if not workbook.sheetnames:
            print(f"Ошибка: Нет листов в файле '{input_path}'.")
            return None, None
        worksheet = workbook[workbook.sheetnames[0]]

        for row_num, row_cells_tuple in enumerate(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row), start=2):
            row_cells = list(row_cells_tuple)

            non_empty_info = [(idx, cell.coordinate, cell.value)
                              for idx, cell in enumerate(row_cells)
                              if not is_likely_empty(cell.value)]
            if not non_empty_info:
                continue

            cell_A = row_cells[start_id_col_idx] if len(row_cells) > start_id_col_idx else None
            cell_I = row_cells[price_total_col_idx] if len(row_cells) > price_total_col_idx else None
            cell_J = row_cells[item_individual_price_col_idx] if len(row_cells) > item_individual_price_col_idx else None
            cell_A_value_str = str(cell_A.value).strip() if cell_A and not is_likely_empty(cell_A.value) else ""

            row_type = None
            
            header_merge_AK = check_merge(worksheet, row_num, 0, 10) 
            footer_text_merge_AH = check_merge(worksheet, row_num, 0, 7) 
            footer_price_merge_IJ = check_merge(worksheet, row_num, price_total_col_idx, item_individual_price_col_idx)

            if header_merge_AK:
                 if cell_A_value_str.startswith("Раздел"): row_type = "section_header"
                 elif cell_A_value_str.startswith("Подраздел"): row_type = "subsection_header"
            
            if row_type is None: 
                if footer_text_merge_AH and footer_price_merge_IJ:
                    if cell_A_value_str.startswith("Итого по подразделу"): row_type = "subsection_footer"
                    elif cell_A_value_str.startswith("Итого по разделу"): row_type = "section_footer"

            if row_type is None:
                non_empty_cell_indices = {info[0] for info in non_empty_info}
                cond1 = non_empty_cell_indices.issuperset(price_row_must_be_non_empty_indices)
                cond2 = non_empty_cell_indices.isdisjoint(price_row_must_be_empty_indices)
                if cond1 and cond2: row_type = "item_price_row"
                else:
                    if cell_A and cell_A.data_type != 'f' and not is_likely_empty(cell_A.value):
                         try:
                             float(str(cell_A.value).replace(',', '.').strip())
                             row_type = "item"
                         except (ValueError, TypeError): pass

            if row_type in ["section_header", "subsection_header", "section_footer", "subsection_footer"] and active_items_buffer:
                if first_section_found: processed_rows_list.extend(active_items_buffer)
                active_items_buffer = []

            if row_type == "section_header":
                if first_section_found:
                    if pending_subsection_header: processed_rows_list.append(pending_subsection_header)
                    if pending_section_header: processed_rows_list.append(pending_section_header)
                pending_section_header = {
                    "type": "header", "level": "section", "start_row": row_num,
                    "col_1_coord": header_merge_AK, "col_3_value": cell_A_value_str,
                    "col_6_value": None, "col_6_coord": None
                }
                pending_subsection_header = None
                first_section_found = True
            elif row_type == "subsection_header":
                if first_section_found and pending_subsection_header:
                    processed_rows_list.append(pending_subsection_header)
                pending_subsection_header = {
                    "type": "header", "level": "subsection", "start_row": row_num,
                    "col_1_coord": header_merge_AK, "col_3_value": cell_A_value_str,
                    "col_6_value": None, "col_6_coord": None
                }
            elif row_type == "subsection_footer":
                if pending_subsection_header:
                    pending_subsection_header["col_6_value"] = cell_I.value if cell_I else None
                    pending_subsection_header["col_6_coord"] = footer_price_merge_IJ
                    if first_section_found:
                        processed_rows_list.append(pending_subsection_header)
                        pending_subsection_header = None
                elif first_section_found:
                    print(f"  [WARN] Строка {row_num}: Итого по подразделу найдено, но не было активного подраздела.")
            elif row_type == "section_footer":
                if first_section_found and pending_subsection_header:
                    processed_rows_list.append(pending_subsection_header)
                    pending_subsection_header = None
                if pending_section_header:
                    pending_section_header["col_6_value"] = cell_I.value if cell_I else None
                    pending_section_header["col_6_coord"] = footer_price_merge_IJ
                    if first_section_found:
                        processed_rows_list.append(pending_section_header)
                        pending_section_header = None
                elif first_section_found:
                    print(f"  [WARN] Строка {row_num}: Итого по разделу найдено, но не было активного раздела.")
            elif row_type == "item_price_row":
                price_total_value = cell_I.value if cell_I else None
                price_total_coord = cell_I.coordinate if cell_I else None
                if active_items_buffer:
                    for item_in_buffer in active_items_buffer: # item переименован во избежание конфликта
                        item_in_buffer["col_6_value"] = price_total_value
                        item_in_buffer["col_6_coord"] = price_total_coord
                    if first_section_found:
                        processed_rows_list.extend(active_items_buffer)
                    active_items_buffer = []
            elif row_type == "item": 
                if first_section_found:
                    cell_J_value = cell_J.value if cell_J else None
                    if is_zero(cell_J_value): # Пропуск если J=0 (для всех item: и материалов, и основных)
                        skipped_items_zero_j_count += 1
                    else:
                        item_data = {
                            "type": "item", "start_row": row_num,
                            "col_6_value": None, "col_6_coord": None
                        }
                        for i in range(min(5, len(row_cells))):
                            cell = row_cells[i]
                            item_data[f"col_{i+1}_value"] = cell.value
                            item_data[f"col_{i+1}_coord"] = cell.coordinate
                        
                        item_id_nature = get_item_id_nature(cell_A.value)
                        if item_id_nature == "decimal": # Это материал
                            # --- ИЗМЕНЕНИЕ ЗДЕСЬ ---
                            # Цена для материала берется из колонки J ЭТОЙ ЖЕ строки
                            # cell_J уже получен и проверен на is_zero выше
                            if cell_J: # Если cell_J существует (и не ноль)
                                item_data["col_6_value"] = cell_J.value
                                item_data["col_6_coord"] = cell_J.coordinate
                            # --- КОНЕЦ ИЗМЕНЕНИЯ ---
                            processed_rows_list.append(item_data)
                        elif item_id_nature == "integer": # Это основная работа/позиция
                            active_items_buffer.append(item_data) # Ожидает общую цену из item_price_row
                        # else: # not_a_number, пропускаем

        if first_section_found:
            if active_items_buffer: processed_rows_list.extend(active_items_buffer)
            if pending_subsection_header: processed_rows_list.append(pending_subsection_header)
            if pending_section_header: processed_rows_list.append(pending_section_header)

        processed_rows_list.sort(key=lambda x: x.get('start_row', float('inf')))
        all_coords_data = []
        skipped_final_price_count = 0
        for row_data in processed_rows_list:
            total_cost_value = row_data.get("col_6_value")
            # Фильтр нулевой итоговой цены ТОЛЬКО для item'ов (не для заголовков/футеров)
            if is_zero(total_cost_value) and row_data.get("type") == "item":
                skipped_final_price_count += 1
                continue

            coords_row = [None] * len(output_headers)
            item_type = row_data.get("type")

            if item_type == "header":
                coords_row[0] = get_start_coord(row_data.get("col_1_coord"))
                coords_row[2] = row_data.get("col_3_value")
                coords_row[5] = get_start_coord(row_data.get("col_6_coord"))
            elif item_type == "item":
                for i in range(min(5, len(output_headers))):
                    coords_row[i] = get_start_coord(row_data.get(f"col_{i+1}_coord"))
                if len(output_headers) > 5:
                     coords_row[5] = get_start_coord(row_data.get("col_6_coord"))
            
            all_coords_data.append(coords_row)
            
        # print(f"Пропущено позиций с нулевой ценой в J: {skipped_items_zero_j_count}")
        # print(f"Пропущено item'ов с нулевой итоговой ценой: {skipped_final_price_count}")
        return output_headers, all_coords_data

    except FileNotFoundError:
        print(f"[ОШИБКА] Файл не найден: {input_path}")
        return None, None
    except Exception as e:
        print(f"[КРИТИЧЕСКАЯ ОШИБКА] при обработке файла '{input_path}' (Смета ру): {e}")
        traceback.print_exc()
        return None, None
    finally:
        if workbook:
            try: workbook.close()
            except Exception as close_e: print(f"  [WARN] Не удалось закрыть Excel файл '{input_path}': {close_e}")