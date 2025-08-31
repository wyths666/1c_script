import pandas as pd
from datetime import datetime
import os

# === 1. Проверка файлов ===
files = ['ostatki.xlsx', 'prodano.xlsx']
for f in files:
    if not os.path.exists(f):
        print(f"Файл {f} не найден!")
        exit()

# === 2. Загружаем остатки ===
try:
    df_ost = pd.read_excel('ostatki.xlsx', engine='openpyxl')
except Exception as e:
    print(f"Ошибка чтения остатков: {e}")
    exit()

df_ost = df_ost.dropna(how='all')
df_ost.columns = ['', 'Номенклатура', 'Остаток_в_магазине', 'Остаток_на_складе']

# Обработка чисел
df_ost['Остаток_в_магазине'] = pd.to_numeric(df_ost['Остаток_в_магазине'], errors='coerce').fillna(0)
df_ost['Остаток_на_складе'] = pd.to_numeric(df_ost['Остаток_на_складе'], errors='coerce').fillna(0)

# Обрезаем последние 4 символа
df_ost['Номенклатура'] = df_ost['Номенклатура'].apply(
    lambda x: x[:-4] if isinstance(x, str) and len(x) > 4 else x
)

# === 3. Загружаем продажи ===
try:
    df_prod = pd.read_excel('prodano.xlsx', engine='openpyxl')
except Exception as e:
    print(f"Ошибка чтения продаж: {e}")
    exit()

df_prod = df_prod.dropna(how='all')
df_prod.columns = ['', 'Номенклатура', 'Продано']
df_prod['Продано'] = pd.to_numeric(df_prod['Продано'], errors='coerce').fillna(0)

# Обрезаем номенклатуру
df_prod['Номенклатура'] = df_prod['Номенклатура'].apply(
    lambda x: x[:-4] if isinstance(x, str) and len(x) > 4 else x
)

# === 4. Объединяем данные ===
df = pd.merge(df_ost, df_prod, on='Номенклатура', how='outer')
df['Продано'] = df['Продано'].fillna(0)
df['Остаток_в_магазине'] = df['Остаток_в_магазине'].fillna(0)
df['Остаток_на_складе'] = df['Остаток_на_складе'].fillna(0)

# === 🔹 ПОТОК 1: Товары, которых нет в магазине, но есть на складе ===
# Даже если продаж нет
df1 = df[
    (df['Остаток_в_магазине'] == 0) &
    (df['Остаток_на_складе'] > 0)
].copy()

# Рекомендация = всё, что есть на складе
df1['Рекомендуется к заказу'] = 1

# Оставляем нужные столбцы
df1 = df1[[
    'Номенклатура',
    'Продано',
    'Остаток_в_магазине',
    'Остаток_на_складе',
    'Рекомендуется к заказу'
]]

# === 🔹 ПОТОК 2: Товары по продажам (где нужно дозаказать) ===
# Только если были продажи
df2 = df[df['Продано'] > 0].copy()

# Рассчитываем рекомендацию
df2['Рекомендуется к заказу'] = df2['Продано'] - df2['Остаток_в_магазине']
df2['Рекомендуется к заказу'] = df2['Рекомендуется к заказу'].clip(lower=0)  # не меньше 0

# Ограничиваем по остатку на складе
df2['Рекомендуется к заказу'] = df2[['Рекомендуется к заказу', 'Остаток_на_складе']].min(axis=1)

# Фильтруем: только где рекомендация > 0
df2 = df2[df2['Рекомендуется к заказу'] > 0].copy()

# Оставляем нужные столбцы
df2 = df2[[
    'Номенклатура',
    'Продано',
    'Остаток_в_магазине',
    'Остаток_на_складе',
    'Рекомендуется к заказу'
]]

# === 5. Объединяем два потока ===
result = pd.concat([df1, df2], ignore_index=True)

# Убираем дубликаты по номенклатуре (если товар попал в оба списка — оставляем один)
result = result.drop_duplicates(subset=['Номенклатура'], keep='first')

# Переименовываем столбцы
result.columns = [
    'Номенклатура',
    'Продажи за 3 недели',
    'Остаток в магазине',
    'Остаток на складе',
    'Рекомендуется к заказу'
]

# Сортируем: сначала те, где рекомендация >0
result = result.sort_values(by='Рекомендуется к заказу', ascending=False)

# Сохраняем
current_date = datetime.now().strftime('%d-%m-%Y')
result.to_excel(f'анализ продаж и остатков от {current_date}.xlsx', index=False)

print(f"Готово! Сформировано {len(result)} позиций.")
print(f"Результат сохранён в 'анализ продаж и остатков от {current_date}.xlsx'")