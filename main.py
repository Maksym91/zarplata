import pandas as pd
import openpyxl
from datetime import datetime

print("=" * 60)
print("СКРИПТ ОБРОБКИ ПРОЦЕНТІВКИ ТА ВІДОМОСТІ")
print("=" * 60)

# Функція для очищення назв колонок
def clean_column_name(col):
    if pd.isna(col):
        return ''
    return str(col).strip().replace('\xa0', ' ')

# --- Зчитування процентівки з 7-го рядка (header=6) ---
print("\n1. Зчитування процентівки...")
procentivka = pd.read_excel('Процентівка.xlsx', header=6)
procentivka.columns = [clean_column_name(col) for col in procentivka.columns]

# Знаходимо колонку з прізвищами
name_col = None
for col in procentivka.columns:
    if isinstance(col, str) and ('прізвище' in col.lower() or 'ініціали' in col.lower()):
        name_col = col
        break

if name_col is None:
    raise ValueError("❌ Не знайдено колонку з прізвищами та ініціалами у процентівці")

print(f"✓ Знайдено колонку з прізвищами: '{name_col}'")
print(f"  Кількість записів у процентівці: {len(procentivka)}")

# Визначаємо колонки для переносу
required_cols = {
    'Дата з': None,
    'Дата по': None,
    'Тарифний розряд': None,
    'розмір премії у відсотках (Р-16 від 10.02.23)': None
}

for col in procentivka.columns:
    col_clean = col.strip().lower()
    if 'дата з' in col_clean:
        required_cols['Дата з'] = col
    elif 'дата по' in col_clean:
        required_cols['Дата по'] = col
    elif 'тарифний розряд' in col_clean:
        required_cols['Тарифний розряд'] = col
    elif 'розмір премії' in col_clean and 'р-16' in col_clean:
        required_cols['розмір премії у відсотках (Р-16 від 10.02.23)'] = col

print("\n✓ Колонки для переносу:")
for key, value in required_cols.items():
    print(f"  - {key}: {value if value else 'НЕ ЗНАЙДЕНО'}")

# --- Зчитування відомості з листа "ПЕРЕНЕСЕННЯ ДАНИХ" ---
print("\n2. Зчитування відомості...")
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
print(f"  Кількість записів у відомості: {len(vidomist)}")

# Створюємо нові колонки у відомості якщо їх немає
for col_name in required_cols.values():
    if col_name and col_name not in vidomist.columns:
        vidomist[col_name] = None

# Створюємо копію для оновлення
vidomist_updated = vidomist.copy()

# --- Обробка даних ---
print("\n3. Обробка даних...")

# Групуємо дані з процентівки за прізвищами (для виявлення дублікатів)
procentivka_by_name = {}
for idx, row in procentivka.iterrows():
    name = row.get(name_col)
    if pd.isna(name) or str(name).strip() == '':
        continue
    
    name_clean = str(name).strip()
    if name_clean not in procentivka_by_name:
        procentivka_by_name[name_clean] = []
    procentivka_by_name[name_clean].append((idx, row))

# Виявляємо дублікати
duplicates_in_procentivka = {name: rows for name, rows in procentivka_by_name.items() if len(rows) > 1}
if duplicates_in_procentivka:
    print(f"  ⚠ Виявлено дублікати у процентівці ({len(duplicates_in_procentivka)}):")
    for name, rows in list(duplicates_in_procentivka.items())[:5]:
        print(f"    - {name}: {len(rows)} дублікатів")

rows_to_add = []  # Дублікати для додавання у відомість
missing_records = []  # Записи яких немає у відомості
updated_count = 0
added_duplicates_count = 0

# Для записів з невідповідностями
mismatch_records = []
real_mismatches = []

# Обробка кожного унікального імені
for name, rows in procentivka_by_name.items():
    matches = vidomist_updated[vidomist_updated[vidomist_name_col] == name]

    if not matches.empty:
        # Отримуємо дані з процентівки
        idx, row = rows[0]
        
        # Оновлюємо дані у всіх існуючих збігах
        for match_idx in matches.index:
            # Перевіряємо розбіжності даних
            has_mismatch = False
            differences = []
            
            # Порівнюємо кожну колонку
            for col_name, orig_col in required_cols.items():
                if orig_col:
                    procentivka_value = row.get(orig_col)
                    vidomist_value = vidomist_updated.at[match_idx, orig_col] if orig_col in vidomist_updated.columns else None
                    
                    # Нормалізуємо значення для порівняння
                    procentivka_str = str(procentivka_value).strip() if pd.notna(procentivka_value) else ''
                    vidomist_str = str(vidomist_value).strip() if pd.notna(vidomist_value) else ''
                    
                    # Порівнюємо тільки якщо обидва значення не порожні
                    if procentivka_str and vidomist_str:
                        if procentivka_str != vidomist_str:
                            has_mismatch = True
                            differences.append(f"{col_name}: '{vidomist_str}' ≠ '{procentivka_str}'")
                    elif (procentivka_str and not vidomist_str) or (not procentivka_str and vidomist_str):
                        has_mismatch = True
                        differences.append(f"{col_name}: порожнє значення ≠ '{procentivka_str if procentivka_str else vidomist_str}'")
            
            # Якщо є невідповідності, зберігаємо детальну інформацію
            if has_mismatch:
                mismatch_info = {
                    vidomist_name_col: name,
                    'Посада (відомість)': vidomist_updated.at[match_idx, 'Посада'] if 'Посада' in vidomist_updated.columns else '',
                    'Посада (процентівка)': row.get('Посада', ''),
                    'Військове звання (відомість)': vidomist_updated.at[match_idx, 'Військове звання'] if 'Військове звання' in vidomist_updated.columns else '',
                    'Військове звання (процентівка)': row.get('Військове звання', ''),
                    'Розбіжності': ' | '.join(differences)
                }
                
                # Додаємо конкретні колонки з розбіжностями
                for col_name, orig_col in required_cols.items():
                    if orig_col:
                        procentivka_value = row.get(orig_col)
                        vidomist_value = vidomist_updated.at[match_idx, orig_col] if orig_col in vidomist_updated.columns else None
                        mismatch_info[f'{col_name} (відомість)'] = vidomist_value if pd.notna(vidomist_value) else ''
                        mismatch_info[f'{col_name} (процентівка)'] = procentivka_value if pd.notna(procentivka_value) else ''
                
                mismatch_records.append(mismatch_info)
                print(f"  ⚠ РОЗБІЖНІСТЬ: '{name}' - {len(differences)} відмінностей")
                for diff in differences[:2]:  # Показуємо перші 2 розбіжності
                    print(f"     • {diff}")
            
            # Оновлюємо дані з процентівки
            for col_name, orig_col in required_cols.items():
                if orig_col and not pd.isna(row.get(orig_col)):
                    vidomist_updated.at[match_idx, orig_col] = row.get(orig_col)
            updated_count += 1
            
            # Якщо є додаткові записи у процентівці, додаємо їх як нові рядки ПІД поточним
            if len(rows) > 1:
                for extra_idx, extra_row in rows[1:]:
                    new_row = vidomist_updated.loc[match_idx].copy()
                    for col_name, orig_col in required_cols.items():
                        if orig_col and not pd.isna(extra_row.get(orig_col)):
                            new_row[orig_col] = extra_row.get(orig_col)
                    # Зберігаємо індекс для вставки під поточний рядок
                    rows_to_add.append((match_idx, new_row))
                    added_duplicates_count += 1
    else:
        # Запис відсутній у відомості - додаємо до відомості з міткою
        print(f"  ⚠ ПРОБЛЕМА: '{name}' є у процентівці, але відсутній у відомості! Додаємо...")
        for idx, row in rows:
            missing_record = {
                vidomist_name_col: name,
                'Посада': row.get('Посада', ''),
                'Військове звання': row.get('Військове звання', ''),
                '_is_missing': True  # Мітка для виділення
            }
            # Додаємо колонки з процентівки
            for col_name, orig_col in required_cols.items():
                if orig_col:
                    missing_record[col_name] = row.get(orig_col)
                    missing_record[f'{col_name}_SOURCE'] = 'ПРОЦЕНТІВКА (ВІДСУТНІЙ)'
            
            # Додаємо до відомості як новий рядок
            vidomist_updated = pd.concat([vidomist_updated, pd.DataFrame([missing_record])], ignore_index=True)
            missing_records.append(missing_record)

# Додаємо дублікати до відомості (вставляємо під відповідні записи)
if rows_to_add:
    print(f"  → Додаємо {len(rows_to_add)} дублікатів під відповідні записи...")
    
    # Сортуємо за індексом (від більшого до меншого), щоб не зсувати індекси
    rows_to_add_sorted = sorted(rows_to_add, key=lambda x: x[0], reverse=True)
    
    # Додаємо рядки по одному під відповідні індекси
    for match_idx, new_row in rows_to_add_sorted:
        # Знаходимо позицію в vidomist_updated для match_idx
        vidomist_idx = vidomist_updated.index.get_loc(match_idx)
        # Додаємо новий рядок після vidomist_idx
        vidomist_updated = pd.concat([
            vidomist_updated.iloc[:vidomist_idx + 1],
            pd.DataFrame([new_row]),
            vidomist_updated.iloc[vidomist_idx + 1:]
        ]).reset_index(drop=True)

# Знаходимо записи які є у відомості, але відсутні у процентівці
print("\n5. Перевірка відсутніх записів...")
vidomist_names = set(vidomist[vidomist_name_col].dropna().astype(str).str.strip())
procentivka_names = set([name for name in procentivka_by_name.keys()])

missing_in_procentivka = []
for name in vidomist_names:
    if name not in procentivka_names and name != '' and name != 'nan':
        print(f"  ⚠ ПРОБЛЕМА: '{name}' є у відомості, але відсутній у процентівці!")
        # Знаходимо всі рядки з цим ім'ям у відомості
        matching_rows = vidomist[vidomist[vidomist_name_col] == name]
        for idx, row in matching_rows.iterrows():
            missing_in_procentivka.append(row.to_dict())

# Підсумки перевірок
print(f"\n📊 РЕЗУЛЬТАТИ ПЕРЕВІРКИ:")
print(f"  - Значень у процентівці: {len(procentivka_by_name)}")
print(f"  - Значень у відомості: {len(vidomist_names)}")
print(f"  - Відсутні у відомості: {len(missing_records)}")
print(f"  - Відсутні у процентівці: {len(missing_in_procentivka)}")

# --- Збереження результатів ---
print("\n6. Збереження результатів...")

# Зберігаємо оновлену відомість з унікальною назвою
from datetime import datetime
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = f'Відомість_оновлена_{timestamp}.xlsx'
print(f"  → Зберігаємо: {output_file}")

# Зберігаємо оновлену відомість
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    vidomist_updated.to_excel(writer, sheet_name='ПЕРЕНЕСЕННЯ ДАНИХ', index=False, startrow=6, header=False)

print(f"✓ Файл збережено: {output_file}")

# Створюємо окремий файл з розбіжностями (тільки справжні розбіжності, не просто порожні значення)
real_mismatches = []
for mismatch in mismatch_records:
    differences = mismatch.get('Розбіжності', '')
    # Перевіряємо чи є реальні розбіжності (не просто порожні значення)
    if '≠' in differences and not all('порожнє значення' in d for d in differences.split('|')):
        real_mismatches.append(mismatch)

if real_mismatches:
    mismatch_df = pd.DataFrame(real_mismatches)
    mismatch_file = 'Розбіжності_за_прізвищем.xlsx'
    print(f"\n  → Створюємо файл з розбіжностями: {mismatch_file}")
    print(f"    Кількість справжніх розбіжностей: {len(real_mismatches)}")
    mismatch_df.to_excel(mismatch_file, index=False)
    print(f"✓ Файл збережено: {mismatch_file}")
else:
    print("\n  ✓ Справжніх розбіжностей не виявлено.")

# Створюємо файл з відсутніми записами у процентівці
if missing_in_procentivka:
    missing_in_procentivka_df = pd.DataFrame(missing_in_procentivka)
    missing_in_procentivka_file = 'Відсутні_у_процентівці.xlsx'
    print(f"\n  → Створюємо файл з відсутніми записами у процентівці: {missing_in_procentivka_file}")
    print(f"    Кількість відсутніх записів у процентівці: {len(missing_in_procentivka)}")
    missing_in_procentivka_df.to_excel(missing_in_procentivka_file, index=False)
    print(f"✓ Файл збережено: {missing_in_procentivka_file}")
else:
    print("\n  ✓ Всі записи з відомості є у процентівці.")

# --- Підсумки ---
print("\n" + "=" * 60)
print("РЕЗУЛЬТАТИ ОБРОБКИ")
print("=" * 60)
print(f"✓ Оброблено записів у відомості: {len(vidomist_updated)}")
print(f"✓ Оновлено записів з процентівки: {updated_count}")
print(f"✓ Додано дублікатів під записами: {added_duplicates_count}")
print(f"✓ Додано відсутніх записів у відомість: {len(missing_records)}")
print(f"✓ Відсутніх у процентівці: {len(missing_in_procentivka)}")
print(f"✓ Справжніх розбіжностей даних: {len(real_mismatches)}")
print("=" * 60)
print("\n✅ Обробка завершена успішно!")
