import asyncio
from pathlib import Path
import pandas as pd
import numpy as np
import time
import re
from datetime import datetime
from utils.convert_style import redactor
from utils.recover_files import convert_with_excel


async def prihod():
    global df
    start_time = time.perf_counter()
    current_date = datetime.now().strftime('%d-%m-%Y')
    output_dir = Path('C:/MyProjects/1c_scripts/отчеты')
    output_dir.mkdir(exist_ok=True)
    file = Path('C:/MyProjects/1c_scripts/остатки') / 'приход.xlsx'
    file_2 = Path('C:/MyProjects/1c_scripts/остатки') / 'sales.xlsx'
    try:
        df = pd.read_excel(file, skiprows=10, engine='openpyxl')
    except Exception as e:
        await convert_with_excel(file, file)
        df = pd.read_excel(file, skiprows=10, engine='openpyxl')
    dlina = len(df)
    print(f'открыт файл с приходом на {dlina} позиций')
    names = ['Маркса', 'Склад']
    df.columns = ['', 'Номенклатура'] + names
    df = df.drop('', axis=1)
    df.iloc[:, 1:] = df.iloc[:, 1:].apply(pd.to_numeric, errors='coerce').fillna(0)  # заменяем Nan на 0


    try:
        df2 = pd.read_excel(file_2, skiprows=8, engine='openpyxl')
        print(f'открыт файл с продажами')
        df2.columns = ['', 'Номенклатура', 'Продажи']
        df2 = df2.drop('', axis=1)
    except FileNotFoundError:
        df2 = pd.DataFrame(columns=['Номенклатура', 'Продажи'])
        print(f'отсутствует файл с продажами')
    except Exception as e:
        await convert_with_excel(file_2, file_2)
        df2 = pd.read_excel(file_2, skiprows=10, engine='openpyxl')
        print(f'открыт файл с продажами')
        df2.columns = ['', 'Номенклатура', 'Продажи']
        df2 = df2.drop('', axis=1)

    df = pd.merge(df, df2, on='Номенклатура', how='left')
    df['Продажи'] = pd.to_numeric(df['Продажи'], errors='coerce').fillna(0)

    df = df.reindex(columns=['Номенклатура', 'Продажи', "Маркса", 'Склад'])

    conditions = [
        df['Продажи'] - df['Маркса'] > 0,
        df['Продажи'] - df['Маркса'] < 0,
        df['Продажи'] - df['Маркса'] == 0
    ]

    choices = [
        df['Продажи'] - df['Маркса'],  # если положительное
        0,  # если отрицательное
        1  # если ноль
    ]

    df.insert(2, 'Рекомендовано к заказу', np.select(conditions, choices, 0))
    df['ordered'] = False
    df = df[(df['Рекомендовано к заказу'] > 0)]

    for idx in df.index:
        if df.loc[idx, "Склад"] >= df.loc[idx, "Рекомендовано к заказу"]:
            df.loc[idx, 'ordered'] = True
        elif 0 < df.loc[idx, "Склад"] < df.loc[idx, "Рекомендовано к заказу"]:
            df.loc[idx, "Рекомендовано к заказу"] = df.loc[idx, "Склад"]
            df.loc[idx, 'ordered'] = True

    result = df[(df['ordered'] == True)]
    output_file = output_dir / f'Склад приход от {current_date}.xlsx'
    result[['Номенклатура', "Рекомендовано к заказу", 'Маркса', 'Склад']].to_excel(output_file, index=False)
    redactor(output_file)
    print(f"создан файл 'заказы со склада (приход) от {current_date}.xlsx' найдено {len(result)} позиций")



    end_time = time.perf_counter()
    execution_time = end_time - start_time
    print(f'Отчет завершен, обработано {dlina} позиций за {execution_time:.4f} секунд')

if __name__ == "__main__":
    asyncio.run(prihod())
