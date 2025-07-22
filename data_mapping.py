import os
import re
import pandas as pd
import numpy as np
from pathlib import Path
import openpyxl # Импортируем явно для доступа к workbook

# --- КОНФИГУРАЦИЯ ---
INPUT_DIR = "balances"
OUTPUT_FILE = "for_further_analysis.xlsx"
RAW_CODE_DATASET = "raw_code_dataset.txt"

# Коды для специальной обработки, которые выносятся на отдельный лист
SPECIAL_RAW_CODES = {
    '3231', '3232', '3233', '3234', '3235', '3236', '3237', '3238', '3239',
    '3241', '3242', '3243', '3244', '3245', '3246', '3247', '3248', '3249',
    '3331', '3332', '3333', '3334', '3335', '3336', '3337', '3338', '3339',
    '3341', '3342', '3343', '3344', '3345', '3346', '3347', '3348', '3349'
}

# Ключевые слова для поиска информации об организации
ORG_INFO_KEYWORDS = {
    'name': 'Полное наименование юридического лица',
    'inn': 'ИНН',
    'address': 'Местонахождение (адрес)',
    'unit': 'Единица измерения'
}

# --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

def load_raw_code_lookup(filepath):
    """Загружает справочник кодов из файла raw_code_dataset.txt."""
    lookup = {}
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            for line in f:
                parts = line.strip().split('|')
                if len(parts) >= 2: # Изменено для гибкости
                    name = parts[0].strip()
                    # Проверяем, что есть и третий элемент, и он - число
                    if len(parts) >= 3 and parts[2].strip().isdigit():
                        code = parts[2].strip()
                        lookup[name.lower()] = code
    except FileNotFoundError:
        print(f"Ошибка: Файл справочника '{filepath}' не найден.")
    return lookup

def clean_value(value):
    """Очищает и преобразует значение ячейки в число."""
    if value is None or isinstance(value, str) and value.strip() in ['-', 'Х', 'x']:
        return np.nan
    if isinstance(value, str):
        # Удаляем скобки, пробелы и заменяем запятую на точку
        value = value.replace('(', '-').replace(')', '').replace(' ', '').replace(',', '.')
    try:
        return float(value)
    except (ValueError, TypeError):
        return np.nan

def find_org_info(sheet):
    """
    Усовершенствованная функция для извлечения информации об организации.
    Ищет ключевое слово и берет значение из первой непустой ячейки справа в той же строке.
    """
    info = {key: None for key in ORG_INFO_KEYWORDS.values()}

    for row in sheet.iter_rows():
        row_values = [cell.value for cell in row]
        for cell_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str):
                cell_text = cell_value.strip()
                for key_name, keyword in ORG_INFO_KEYWORDS.items():
                    if cell_text.startswith(keyword):
                        # Ищем значение в ячейках справа от ключевого слова
                        for next_cell_value in row_values[cell_idx + 1:]:
                            if next_cell_value is not None and str(next_cell_value).strip() != '':
                                info[key_name] = str(next_cell_value).strip()
                                break # Нашли значение, переходим к следующей строке
                        break # Нашли ключевое слово, переходим к следующей строке
    
    # Постобработка для извлечения чистого ИНН
    if info.get('inn'):
        # Используем re.search для поиска 10 или 12 цифр подряд
        match = re.search(r'\b(\d{10}|\d{12})\b', info['inn'])
        if match:
            info['inn'] = match.group(1)
        else:
             # Если регулярное выражение не сработало, применяем простую очистку
             info['inn'] = re.sub(r'\D', '', info['inn'])
             
    return info


def parse_financial_sheet(sheet, inn, raw_code_lookup):
    """Парсит лист с финансовыми данными (Баланс, Отчет о фин. результатах и т.д.)."""
    data_rows = []
    additional_rows = []
    
    header_map = {}
    year_map = {}
    data_start_row = -1

    # 1. Находим заголовок таблицы и определяем столбцы
    for r_idx, row in enumerate(sheet.iter_rows(max_row=30), start=1):
        row_values = [str(cell.value).lower().strip() if cell.value is not None else '' for cell in row]
        
        # Более гибкие условия поиска заголовка
        is_name_present = any('наименование показателя' in v for v in row_values)
        is_code_present = any('код' in v for v in row_values) # Ищем "код", а не "код строки"

        if is_name_present and is_code_present:
            data_start_row = r_idx + 1
            for c_idx, value in enumerate(row_values):
                if 'наименование показателя' in value:
                    header_map['name'] = c_idx
                elif 'код' in value and 'окуд' not in value and 'окпо' not in value: # Исключаем коды ОКУД/ОКПО
                    header_map['code'] = c_idx
                
                # Ищем годы в заголовках столбцов
                year_match = re.search(r'\b(20\d{2})\b', value)
                if year_match:
                    year = int(year_match.group(1))
                    year_map[year] = c_idx
            break

    if not header_map or data_start_row == -1:
        # print(f"  - Не удалось найти заголовок таблицы на листе '{sheet.title}'")
        return [], []

    sorted_years = sorted(year_map.keys(), reverse=True)
    
    last_code = None
    last_code_row_idx = -1
    
    # 2. Итерируемся по строкам с данными
    for r_idx, row in enumerate(sheet.iter_rows(min_row=data_start_row), start=data_start_row):
        name_cell_val = row[header_map.get('name')].value if 'name' in header_map else None
        code_cell_val = row[header_map.get('code')].value if 'code' in header_map else None
        
        indicator_name = str(name_cell_val).strip() if name_cell_val else None
        if not indicator_name or indicator_name.isdigit(): # Пропускаем строки без наименования
            continue

        raw_code = str(code_cell_val).strip() if code_cell_val and str(code_cell_val).isdigit() else None
        
        # Логика определения raw_code
        if not raw_code:
            # Сначала ищем по точному совпадению
            if indicator_name.lower() in raw_code_lookup:
                raw_code = raw_code_lookup[indicator_name.lower()]
            else: # Если не нашли, ищем по частичному совпадению
                for name, code in raw_code_lookup.items():
                    if name in indicator_name.lower():
                        raw_code = code
                        break

            if not raw_code and last_code:
                row_diff = r_idx - last_code_row_idx
                raw_code = f"{last_code}.{row_diff:02d}"

        if raw_code:
            if '.' not in raw_code:
                last_code = raw_code
                last_code_row_idx = r_idx

            for year_val in sorted_years:
                col_idx = year_map[year_val]
                saldo = clean_value(row[col_idx].value)
                
                if pd.notna(saldo):
                    record = {
                        'ИНН': inn,
                        'Год': year_val,
                        'Наименование показателя': indicator_name,
                        'raw_code': raw_code,
                        'Значение': saldo
                    }
                    if raw_code in SPECIAL_RAW_CODES:
                        additional_rows.append(record)
                    else:
                        data_rows.append(record)
                        
    return data_rows, additional_rows

# --- ОСНОВНАЯ ЛОГИКА ---
def main():
    """Основная функция для запуска процесса."""
    base_path = Path(__file__).parent
    input_path = base_path / INPUT_DIR
    
    if not input_path.exists():
        print(f"Ошибка: Папка '{INPUT_DIR}' не найдена. Создайте ее и поместите туда файлы отчетов.")
        return

    raw_code_lookup = load_raw_code_lookup(base_path / RAW_CODE_DATASET)
    
    all_orgs_info = []
    all_financial_data = []
    all_additional_data = []

    excel_files = list(input_path.glob("*.xlsx"))
    if not excel_files:
        print(f"В папке '{INPUT_DIR}' не найдены файлы .xlsx для обработки.")
        return

    print(f"Найдено {len(excel_files)} файлов для обработки.")

    for file_path in excel_files:
        print(f"Обработка файла: {file_path.name}...")
        try:
            # Используем openpyxl для более надежного чтения
            wb = openpyxl.load_workbook(file_path, data_only=True)
            
            # 1. Извлечение информации об организации
            org_sheet_name = next((s for s in wb.sheetnames if 'сведения об организации' in s.lower()), None)
            if not org_sheet_name:
                print(f"  - Предупреждение: Лист 'Сведения об организации' не найден в {file_path.name}")
                continue
            
            org_sheet = wb[org_sheet_name]
            org_info = find_org_info(org_sheet)
            
            if not org_info.get('inn'):
                 print(f"  - ОШИБКА: Не удалось извлечь ИНН из файла {file_path.name}. Пропускаем.")
                 continue
            
            all_orgs_info.append({
                'Наименование организации': org_info.get('name'),
                'ИНН': org_info.get('inn'),
                'Местонахождение (адрес)': org_info.get('address'),
                'Единица измерения': org_info.get('unit')
            })

            # 2. Извлечение финансовых данных
            for sheet_name in wb.sheetnames:
                if 'сведения об организации' in sheet_name.lower():
                    continue
                
                sheet = wb[sheet_name]
                data, additional = parse_financial_sheet(sheet, org_info['inn'], raw_code_lookup)
                all_financial_data.extend(data)
                all_additional_data.extend(additional)

        except Exception as e:
            print(f"  - КРИТИЧЕСКАЯ ОШИБКА при обработке файла {file_path.name}: {e}")

    if not all_financial_data:
        print("Не удалось извлечь никаких финансовых данных. Выходной файл не будет создан.")
        return

    print("Формирование итогового файла...")
    
    orgs_df = pd.DataFrame(all_orgs_info).drop_duplicates(subset=['ИНН'])
    financial_df = pd.DataFrame(all_financial_data)
    if all_additional_data:
        additional_df = pd.DataFrame(all_additional_data)
    else:
        additional_df = pd.DataFrame(columns=['ИНН', 'Год', 'Наименование показателя', 'raw_code', 'Значение'])


    available_years = sorted(financial_df['Год'].unique(), reverse=True)
    
    final_columns_order = ['ИНН'] + sorted(list(set(financial_df['raw_code'])))

    with pd.ExcelWriter(base_path / OUTPUT_FILE, engine='xlsxwriter') as writer:
        orgs_df.to_excel(writer, sheet_name='Organisations', index=False)
        
        year_labels = ['year', 'year-1', 'year-2']
        for i, year in enumerate(available_years):
            sheet_name = year_labels[i] if i < len(year_labels) else f"year-{i}"
            
            year_df = financial_df[financial_df['Год'] == year]
            if not year_df.empty:
                pivot_df = year_df.pivot_table(
                    index='ИНН', 
                    columns='raw_code', 
                    values='Значение',
                    aggfunc='first'
                ).reset_index()
                
                # Приводим к единому набору колонок для всех листов
                current_cols = pivot_df.columns.tolist()
                all_cols = sorted(list(set(current_cols + final_columns_order)))
                
                for col in all_cols:
                    if col not in pivot_df.columns:
                        pivot_df[col] = np.nan
                
                # Устанавливаем порядок: сначала ИНН, потом остальные по возрастанию
                display_cols = ['ИНН'] + [c for c in all_cols if c != 'ИНН']
                pivot_df = pivot_df[display_cols]
                pivot_df.to_excel(writer, sheet_name=str(year), index=False)
        
        if not additional_df.empty:
            additional_df.to_excel(writer, sheet_name='additional_raw_codes', index=False)

    print(f"Готово! Результаты сохранены в файл: {OUTPUT_FILE}")

if __name__ == '__main__':
    main()

#--- THIS IS WORKING VERSION ---#