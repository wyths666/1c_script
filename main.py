import asyncio
import time
from utils.electronics import electronics
from utils.instrument import instruments
from utils.pults import pults
from utils.radio import radio
from utils.spare_parts import zip_otchet
from utils.batteries import batteries


async def run_module_async(module_func, module_name):
    """Запускает модуль асинхронно"""
    print(f"Запуск модуля: {module_name}")
    start_time = time.perf_counter()

    try:
        # Предполагаем, что функции уже асинхронные
        result = await module_func()
        end_time = time.perf_counter()
        execution_time = end_time - start_time
        print(f"Модуль {module_name} завершен за {execution_time:.2f} секунд")
        return result
    except Exception as e:
        print(f"Ошибка в модуле {module_name}: {e}")
        return None


async def main():
    """Основная асинхронная функция"""
    print("Запуск всех модулей параллельно...")
    start_total = time.perf_counter()

    # Запускаем все модули параллельно
    tasks = [
        run_module_async(electronics, "Электроника"),
        run_module_async(instruments, "Инструменты"),
        run_module_async(pults, "Пульты"),
        run_module_async(radio, "Радио"),
        run_module_async(zip_otchet, "Запчасти"),
        run_module_async(batteries, "Батарейки")
    ]

    # Ждем завершения всех задач
    results = await asyncio.gather(*tasks, return_exceptions=True)

    end_total = time.perf_counter()
    total_time = end_total - start_total
    print(f"\nВсе модули завершены за {total_time:.2f} секунд")

    return results


if __name__ == "__main__":
    asyncio.run(main())