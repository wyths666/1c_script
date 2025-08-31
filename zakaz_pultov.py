import pandas as pd
import os

filename = 'ostatki.xlsx'

# Проверяем наличие файла
if not os.path.exists(filename):
    print(f"Файл {filename} не найден!")
    exit()

try:
    df = pd.read_excel(filename, engine='openpyxl')
except Exception as e:
    print(f"Ошибка чтения файла: {e}")
    exit()

# Удаляем полностью пустые строки
df = df.dropna(how='all')

# Переименовываем столбцы
df.columns = ['', 'Номенклатура', 'Остаток_в_магазине', 'Остаток_на_складе']

# Приводим числовые столбцы к числу, заменяя пустые на 0
df['Остаток_в_магазине'] = pd.to_numeric(df['Остаток_в_магазине'], errors='coerce').fillna(0)
df['Остаток_на_складе'] = pd.to_numeric(df['Остаток_на_складе'], errors='coerce').fillna(0)

# --- НОВОЕ: Обрезаем последние 4 символа у номенклатуры ---
# Проверяем, что значение — строка и длина больше 4
df['Номенклатура'] = df['Номенклатура'].apply(
    lambda x: x[:-4] if isinstance(x, str) and len(x) > 4 else x
)

# Фильтруем: магазин = 0, склад > 0
result = df[(df['Остаток_в_магазине'] == 0) & (df['Остаток_на_складе'] > 1)]

# Сохраняем результат
result[['Номенклатура', 'Остаток_в_магазине', 'Остаток_на_складе']].to_excel('заказ_пульты.xlsx', index=False)

print(f"Готово! Найдено {len(result)} товаров для перевозки.")