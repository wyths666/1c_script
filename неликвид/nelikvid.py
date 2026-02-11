import pandas as pd

from utils.convert_style import redactor

# Пути к файлам
sales_file = 'sales.xlsx'
stock_file = 'stocks.xlsx'

# Читаем файлы
sales_df = pd.read_excel(sales_file)
stock_df = pd.read_excel(stock_file)

# Переименуем столбцы для удобства (если нужно)
sales_df.columns = ['Номенклатура', 'Продажи']
stock_df.columns = ['Номенклатура', 'Остаток']

# Объединяем остатки с продажами
merged_df = stock_df.merge(
    sales_df,
    on='Номенклатура',
    how='left'
)

# Товары, которые не продавались ни разу
not_sold_df = merged_df[
    (merged_df['Продажи'].isna()) | (merged_df['Продажи'] == 0)
]

# Итоговый список
result = not_sold_df[['Номенклатура', 'Остаток']]

# Сохраняем результат
result.to_excel('not_sold_products.xlsx', index=False)
redactor('../utils/not_sold_products.xlsx')
print('Готово. Файл not_sold_products.xlsx создан.')
