import pandas as pd

# загрузка файла
df = pd.read_excel("nomenklatur.xlsx")
df.columns = ["Наименование"]
# чистим пробелы
df["Наименование"] = df["Наименование"].str.strip().str.lower()

# ищем дубли по наименованию
duplicates_name = df[df.duplicated("Наименование", keep=False)]


# сохраняем
duplicates_name.to_excel("duplicates_by_name.xlsx", index=False)


print("Готово 🚀")
