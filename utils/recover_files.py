import win32com.client as win32
import os


def convert_with_excel(input_path, output_path=None):
    """
    Конвертирует файл через автоматизацию Excel
    (требует установленного Microsoft Excel)
    """
    if output_path is None:
        output_path = input_path.replace('.xlsx', '.xlsx')

    try:
        # Запускаем Excel
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False  # Скрываем окно Excel
        excel.DisplayAlerts = False  # Отключаем предупреждения

        # Открываем файл
        workbook = excel.Workbooks.Open(os.path.abspath(input_path))

        # Сохраняем как новый файл (это пересоздает структуру)
        workbook.SaveAs(os.path.abspath(output_path), 51)  # 51 = xlsx

        # Закрываем
        workbook.Close()
        excel.Quit()

        print(f"Файл успешно конвертирован: {output_path}")
        return output_path

    except Exception as e:
        print(f"Ошибка при конвертации через Excel: {e}")
        return None


