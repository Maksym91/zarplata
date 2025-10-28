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

# Обробка кожного унікального імені
for name, rows in procentivka_by_name.items():
    matches = vidomist_updated[vidomist_updated[vidomist_name_col] == name]

    if not matches.empty:
        # Оновлюємо дані у всіх існуючих збігах
        for match_idx in matches.index:
            # Якщо є кілька записів у процентівці, використовуємо перший для оновлення існуючого
            idx, row = rows[0]
            for col_name, orig_col in required_cols.items():
                if orig_col and not pd.isna(row.get(orig_col)):
                    vidomist_updated.at[match_idx, orig_col] = row.get(orig_col)
            updated_count += 1
            
            # Якщо є додаткові записи у процентівці, додаємо їх як нові рядки
            if len(rows) > 1:
                for extra_idx, extra_row in rows[1:]:
                    new_row = vidomist_updated.loc[match_idx].copy()
                    for col_name, orig_col in required_cols.items():
                        if orig_col and not pd.isna(extra_row.get(orig_col)):
                            new_row[orig_col] = extra_row.get(orig_col)
                    rows_to_add.append(new_row)
                    added_duplicates_count += 1
    else:
        # Запис відсутній у відомості
        for idx, row in rows:
            missing_record = {vidomist_name_col: name}
            for col_name, orig_col in required_cols.items():
                if orig_col:
                    missing_record[col_name] = row.get(orig_col)
            missing_records.append(missing_record)

# Додаємо дублікати до відомості
if rows_to_add:
    print(f"  → Додаємо {len(rows_to_add)} дублікатів...")
    new_rows_df = pd.DataFrame(rows_to_add)
    vidomist_updated = pd.concat([vidomist_updated, new_rows_df], ignore_index=True)

# --- Збереження результатів ---
print("\n4. Збереження результатів...")

# Зберігаємо оновлену відомість з унікальною назвою
from datetime import datetime
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = f'Відомість_оновлена_{timestamp}.xlsx'
print(f"  → Зберігаємо: {output_file}")

# Використовуємо ExcelWriter для коректного збереження
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Зберігаємо першу вкладку з оновленими даними
    vidomist_updated.to_excel(writer, sheet_name='ПЕРЕНЕСЕННЯ ДАНИХ', index=False, startrow=6, header=False)
    
    # Якщо потрібно, додамо інші листи з оригінального файлу

print(f"✓ Файл збережено: {output_file}")

# Створюємо файл з відсутніми записами
if missing_records:
    missing_df = pd.DataFrame(missing_records)
    missing_file = 'Відомість_нові.xlsx'
    print(f"\n  → Створюємо файл з відсутніми записами: {missing_file}")
    print(f"    Кількість відсутніх записів: {len(missing_records)}")
    missing_df.to_excel(missing_file, index=False)
    print(f"✓ Файл збережено: {missing_file}")
else:
    print("\n  ✓ Всі записи з процентівки є у відомості.")

# --- Підсумки ---
print("\n" + "=" * 60)
print("РЕЗУЛЬТАТИ ОБРОБКИ")
print("=" * 60)
print(f"✓ Оброблено записів у відомості: {len(vidomist_updated)}")
print(f"✓ Оновлено записів з процентівки: {updated_count}")
print(f"✓ Додано дублікатів: {added_duplicates_count}")
print(f"✓ Відсутніх записів: {len(missing_records)}")
print("=" * 60)
print("\n✅ Обробка завершена успішно!")
