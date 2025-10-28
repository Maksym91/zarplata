import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime

print("=" * 60)
print("СКРИПТ ОБРОБКИ ПРОЦЕНТІВКИ ТА ВІДОМОСТІ")
print("=" * 60)

# Функція для очищення назв колонок
def clean_column_name(col):
    if pd.isna(col):
        return ''
    return str(col).strip().replace('\xa0', ' ')

# --- Зчитування процентівки ---
print("\n1. Зчитування та аналіз процентівки...")
procentivka = pd.read_excel('Процентівка.xlsx', header=6)
procentivka.columns = [clean_column_name(col) for col in procentivka.columns]

# Знаходимо колонку з прізвищами
name_col = None
for col in procentivka.columns:
    if isinstance(col, str) and ('прізвище' in col.lower() and 'ініціали' in col.lower()):
        name_col = col
        break

if name_col is None:
    raise ValueError("❌ Не знайдено колонку з прізвищами та ініціалами у процентівці")

print(f"✓ Знайдено колонку з прізвищами: '{name_col}'")

# Групуємо дублікати за прізвищем
procentivka_duplicates = {}
for idx, row in procentivka.iterrows():
    name = row.get(name_col)
    if pd.notna(name) and str(name).strip():
        name_clean = str(name).strip()
        if name_clean not in procentivka_duplicates:
            procentivka_duplicates[name_clean] = []
        procentivka_duplicates[name_clean].append((idx, row.to_dict()))

# Знаходимо дублікати
duplicates_list = {name: entries for name, entries in procentivka_duplicates.items() if len(entries) > 1}
print(f"  📊 Аналіз процентівки:")
print(f"     - Усього записів: {len(procentivka)}")
print(f"     - Унікальних прізвищ: {len(procentivka_duplicates)}")
print(f"     - Дублікатів: {len(duplicates_list)}")

if duplicates_list:
    print(f"\n  ⚠ Дублікати у процентівці:")
    for name, entries in list(duplicates_list.items())[:5]:
        print(f"     - {name}: {len(entries)} записів")

# --- Зчитування відомості ---
print("\n2. Зчитування та аналіз відомості...")
vidomist = pd.read_excel('Відомість.xlsm', sheet_name='ПЕРЕНЕСЕННЯ ДАНИХ', header=6)
vidomist.columns = [clean_column_name(col) for col in vidomist.columns]

# Знаходимо колонку з прізвищами у відомості
vidomist_name_col = None
for col in vidomist.columns:
    if isinstance(col, str) and ('прізвище' in col.lower() and 'ініціали' in col.lower()):
        vidomist_name_col = col
        break

if vidomist_name_col is None:
    raise ValueError("❌ Не знайдено колонку з прізвищами та ініціалами у відомості")

print(f"✓ Знайдено колонку з прізвищами: '{vidomist_name_col}'")

# Групуємо дублікати у відомості
vidomist_duplicates = {}
for idx, row in vidomist.iterrows():
    name = row.get(vidomist_name_col)
    if pd.notna(name) and str(name).strip():
        name_clean = str(name).strip()
        if name_clean not in vidomist_duplicates:
            vidomist_duplicates[name_clean] = []
        vidomist_duplicates[name_clean].append((idx, row.to_dict()))

print(f"  📊 Аналіз відомості:")
print(f"     - Усього записів: {len(vidomist)}")
print(f"     - Унікальних прізвищ: {len(vidomist_duplicates)}")

# --- Обробка та об'єднання даних ---
print("\n3. Обробка та об'єднання даних...")

result_rows = []
missing_records = []
updated_count = 0
duplicates_added = 0

# Знаходимо відсутні у відомості
vidomist_names_set = set(vidomist_duplicates.keys())
procentivka_names_set = set(procentivka_duplicates.keys())

missing_in_vidomist = []
for name in procentivka_names_set:
    if name not in vidomist_names_set:
        missing_in_vidomist.append(procentivka_duplicates[name][0][1])  # Перший запис

print(f"  - Відсутніх у відомості: {len(missing_in_vidomist)}")

# Обробляємо кожне прізвище з відомості
for name, entries in vidomist_duplicates.items():
    if name in procentivka_names_set:
        # Є у процентівці - беремо дані звідти
        procentivka_entry = procentivka_duplicates[name][0][1]
        
        for idx, vidomist_entry in entries:
            # Об'єднуємо дані
            combined = vidomist_entry.copy()
            
            # Додаємо дані з процентівки
            for col_name in ['Посада', 'Військове звання', 'Дата з', 'Дата по', 'Тарифний розряд', 'розмір премії у відсотках (Р-16 від 10.02.23)']:
                if col_name in procentivka_entry:
                    combined[col_name] = procentivka_entry[col_name]
            
            combined['is_duplicate'] = False
            result_rows.append(combined)
            updated_count += 1
        
        # Додаємо дублікати з процентівки
        if len(procentivka_duplicates[name]) > 1:
            for dup_idx, dup_entry in procentivka_duplicates[name][1:]:
                new_row = entries[0][1].copy()  # Базовий рядок з відомості
                
                # Додаємо дані дубліката з процентівки
                for col_name in ['Посада', 'Військове звання', 'Дата з', 'Дата по', 'Тарифний розряд', 'розмір премії у відсотках (Р-16 від 10.02.23)']:
                    if col_name in dup_entry:
                        new_row[col_name] = dup_entry[col_name]
                
                new_row['is_duplicate'] = True
                result_rows.append(new_row)
                duplicates_added += 1
    else:
        # Немає у процентівці - додаємо як є
        for idx, vidomist_entry in entries:
            vidomist_entry['is_duplicate'] = False
            result_rows.append(vidomist_entry)

# Додаємо відсутні записи
for missing_entry in missing_in_vidomist:
    missing_entry['is_duplicate'] = False
    result_rows.append(missing_entry)

# Створюємо DataFrame
result_df = pd.DataFrame(result_rows)

# --- Збереження з червоним виділенням ---
print("\n4. Збереження результатів...")

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = f'Відомість_оновлена_{timestamp}.xlsx'
print(f"  → Зберігаємо: {output_file}")

# Зберігаємо Excel
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    result_df.to_excel(writer, sheet_name='Результат', index=False)

# Застосовуємо червоне виділення для дублікатів
wb = openpyxl.load_workbook(output_file)
ws = wb['Результат']

red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

# Знаходимо колонку is_duplicate (якщо є)
duplicate_col = None
for col_idx, col_name in enumerate(result_df.columns, 1):
    if 'is_duplicate' in str(col_name):
        duplicate_col = col_idx
        break

if duplicate_col:
    for row_idx in range(2, ws.max_row + 1):
        if ws.cell(row=row_idx, column=duplicate_col).value == True:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col).fill = red_fill

# Прибираємо технічну колонку
if 'is_duplicate' in result_df.columns:
    for row_idx in range(1, ws.max_row + 1):
        ws.cell(row=row_idx, column=duplicate_col).value = None

wb.save(output_file)
print(f"✓ Файл збережено з червоним виділенням дублікатів: {output_file}")

# --- Створення файлу з відсутніми ---
if missing_in_vidomist:
    missing_df = pd.DataFrame(missing_in_vidomist)
    missing_file = 'Відсутні.xlsx'
    print(f"\n  → Створюємо файл з відсутніми: {missing_file}")
    print(f"    Кількість: {len(missing_in_vidomist)}")
    missing_df.to_excel(missing_file, index=False)
    print(f"✓ Файл збережено: {missing_file}")

# --- Підсумки ---
print("\n" + "=" * 60)
print("РЕЗУЛЬТАТИ ОБРОБКИ")
print("=" * 60)
print(f"✓ Усього записів у результаті: {len(result_df)}")
print(f"✓ Оновлено записів: {updated_count}")
print(f"✓ Додано дублікатів (червоні): {duplicates_added}")
print(f"✓ Відсутніх записів: {len(missing_in_vidomist)}")
print("=" * 60)
print("\n✅ Обробка завершена успішно!")
