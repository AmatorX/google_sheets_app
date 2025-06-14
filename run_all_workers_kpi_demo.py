
# import datetime
# import os
# from collections import defaultdict
#
# import django
# import time
#
#
# os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tsa.settings")
# django.setup()
#
# from sheets.tasks import get_worker_materials_hourly_earnings
# from sheets.all_workers_kpi_shhet import WorkersKPIMasterTable
# from sheets.sheet1 import Sheet1Table
# from tsa_app.models import ForemanAndWorkersKPISheet, BuildObject
#
#
#
#
# def run_monthly_summary_tasks():
#     data_dict = get_worker_materials_hourly_earnings(month=5, year=2025)
#     print(f"Для записи дотсутпны данные: {data_dict}")
#     try:
#         write_kpi_data_to_master_table_demo(data_dict)
#     except Exception as e:
#         print(f"Ошибка при сохранении KPI: {e}")
#         raise
#     print("Задача завершена. Данные сохранены в БД и Google Sheets")
#
#
#
# def write_kpi_data_to_master_table_demo(data_dict):
#     """
#     Записывает агрегированные KPI-данные работников в мастер-таблицу Google Sheets.
#     Функция выполняет следующие действия:
#     1. Получает объект ForemanAndWorkersKPISheet за текущий год из БД
#     2. Инициализирует WorkersKPIMasterTable с полученным объектом
#     3. Записывает переданные KPI-данные в соответствующую таблицу Google Sheets
#     Параметры:
#         data_dict (dict): Словарь с данными KPI в формате:
#                          {
#                              "Имя работника1": суммарный_заработок,
#                              "Имя работника2": суммарный_заработок,
#                              ...
#                          }
#     Исключения:
#         ForemanAndWorkersKPISheet.DoesNotExist: Если в базе нет записи за текущий год
#         Exception: Любые другие ошибки при работе с Google Sheets API
#     Пример использования:
#         data = {"Иванов И.": 1500.50, "Петров А.": 2000.00}
#         write_kpi_data_to_master_table(data)
#     """
#     current_year = datetime.datetime.now().year
#     try:
#         # Получаем объект ForemanAndWorkersKPISheet за текущий год
#         kpi_sheet_obj = ForemanAndWorkersKPISheet.objects.get(year=current_year)
#         workers_kpi_master_table = WorkersKPIMasterTable(kpi_sheet_obj)
#         workers_kpi_master_table.write_monthly_kpi(data_dict)
#     except ForemanAndWorkersKPISheet.DoesNotExist:
#         error_msg = f"Не найден объект ForemanAndWorkersKPISheet за {current_year} год."
#         print(error_msg)
#         raise ForemanAndWorkersKPISheet.DoesNotExist(error_msg)
#     except Exception as e:
#         print(f"Ошибка при записи KPI данных в мастер-таблицу: {str(e)}")
#         raise
#
#
# if __name__ == "__main__":
#     run_monthly_summary_tasks()





import os
import django



os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tsa.settings")
django.setup()

from sheets.tasks import run_monthly_summary_tasks


if __name__ == "__main__":
    run_monthly_summary_tasks()





