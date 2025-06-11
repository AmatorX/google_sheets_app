import json
import os
from collections import defaultdict
from datetime import datetime, date, timedelta
from pytz import timezone
from time import sleep

import pytz
import logging
from celery import shared_task

from sheets.all_workers_kpi_shhet import WorkersKPIMasterTable
from sheets.object_kpi_sheet import ObjectKPITable
from sheets.photos_sheet import PhotosTable
from sheets.results_sheet import ResultsTable
from sheets.sheet1 import Sheet1Table
from sheets.users_kpi_sheet import UsersKPITable
from sheets.work_time_sheet import WorkTimeTable
from tsa_app.models import BuildObject, MonthlyKPIData, ForemanAndWorkersKPISheet

logger = logging.getLogger(__name__)

@shared_task
def update_all_worktime_tables():
    logger.info("Запущена задача Celery: обновление WorkTime таблиц")

    build_objects = BuildObject.objects.all()

    for build_object in build_objects:
        try:
            logger.info(f"Обновляем WorkTime таблицу для объекта: {build_object.name}")
            table = WorkTimeTable(obj=build_object)
            table.update_current_chunk()
        except Exception as e:
            logger.error(f"Ошибка при обновлении объекта {build_object.name}: {e}")
        sleep(10)

    logger.info("Задача Celery завершена")

@shared_task
def update_photos_tables():
    logger.info("Celery: запуск задачи обновления таблиц фотоотчётов")

    build_objects = BuildObject.objects.all()

    for build_object in build_objects:
        try:
            logger.info(f"Обрабатываем объект: {build_object.name}")
            photos_table = PhotosTable(obj=build_object)
            photos_table.write_missing_worker_tables()
        except Exception as e:
            logger.error(f"Ошибка при обработке объекта {build_object.name}: {e}")
        sleep(10)

    logger.info("Celery: задача обновления фото-таблиц завершена")


@shared_task
def update_results_tables():
    """
    Celery задача для обновления таблиц результатов в Google Sheets
    для всех строительных объектов.
    """

    build_objects = BuildObject.objects.all()

    for build_object in build_objects:
        try:
            logger.info(f"Обрабатываем объект: {build_object.name}")
            results_table = ResultsTable(obj=build_object)

            # Получаем и строим таблицу для работников
            all_worker_data = results_table.build_worker_table()
            if all_worker_data:
                results_table.write_to_google_sheets(all_worker_data)

        except Exception as e:
            logger.error(f"Ошибка при обработке объекта {build_object.name}: {e}\n")
        sleep(10)
    logger.info("Celery: задача обновления таблиц результатов завершена")


@shared_task
def update_daily_object_kpis():
    build_objects = BuildObject.objects.all()

    if not build_objects.exists():
        print("Нет доступных объектов.")
        return

    for obj in build_objects:
        print(f"\n--- Объект: {obj.name} ---")
        try:
            table = ObjectKPITable(obj)
            table.write_today_data()
            print("Данные успешно записаны.")
        except Exception as e:
            print(f"Ошибка при обработке объекта '{obj.name}': {e}")
        sleep(10)


@shared_task
def update_daily_user_kpis():
    build_objects = BuildObject.objects.all()

    if not build_objects.exists():
        logger.info("Нет доступных объектов.")
        return

    for obj in build_objects:
        logger.info(f"Объект: {obj.name}")
        try:
            table = UsersKPITable(obj)
            table.write_kpi_table()
            logger.info(f"KPI-таблица для пользователей объекта '{obj.name}' успешно записана.")
        except Exception as e:
            logger.error(f"Ошибка при обработке объекта '{obj.name}': {e}")
        sleep(10)

@shared_task
def process_daily_kpi_data_for_tgbot():
    today_str = date.today().strftime("%Y-%m-%d")
    output_dir = "/app/db"
    os.makedirs(output_dir, exist_ok=True)

    filepath = os.path.join(output_dir, f"{today_str}.txt")

    build_objects = BuildObject.objects.all()
    if not build_objects.exists():
        return "Нет доступных объектов."

    result_data = {}

    for obj in build_objects:
        try:
            table = UsersKPITable(obj)
            data = table.get_today_per_day_dict()
            result_data.update(data)
        except Exception as e:
            result_data[f"error_{obj.id}"] = str(e)

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(result_data, f, ensure_ascii=False, indent=2)

    return f"Готово. Данные сохранены в {filepath}"


# def is_last_day_of_month(date_now=None):
#     date_now = date_now or datetime.today()
#     next_day = date_now + timedelta(days=1)
#     return next_day.day == 1
#
# @shared_task
# def run_monthly_summary_tasks():
#     if not is_last_day_of_month():
#         logger.info("Сегодня не последний день месяца. Задача завершена без выполнения.")
#         return
#
#     logger.info("Сегодня последний день месяца. Запуск записи сводных данных...")
#
#     queryset = BuildObject.objects.all()
#     for obj in queryset:
#         try:
#             # 1. KPI по пользователям
#             table = Sheet1Table(obj)
#             table.write_summary_workers()
#             logger.info(f"Сводка результатов работников записана для '{obj.name}'")
#
#             # 2. Сводка по объекту
#             table.write_summary()
#             logger.info(f"Сводная таблица за месяц записана для '{obj.name}'")
#
#         except Exception as e:
#             logger.error(f"Ошибка при обработке '{obj.name}': {e}")
#         sleep(60)
#
#     logger.info("Месячная задача завершена.")


def is_last_day_of_month(date_now=None):
    date_now = date_now or datetime.today()
    next_day = date_now + timedelta(days=1)
    return next_day.day == 1

@shared_task
def run_monthly_summary_tasks():
    """
    Ежедневная задача для сбора и сохранения KPI данных.
    Собирает данные по всем объектам BuildObject и сохраняет:
    - В БД через save_to_db_current_month_kpi()
    - В Google Sheets через write_kpi_data_to_master_table()
    """
    # --- Запуск только в конце месяца (временно отключено) ---
    # edmonton_tz = timezone("America/Edmonton")
    # now_in_edmonton = datetime.now(edmonton_tz)
    # if not is_last_day_of_month(now_in_edmonton):
    #     logger.info("Сегодня не последний день месяца по времени Калгари. Задача завершена без выполнения.")
    #     return
    # ---------------------------------------------------------
    logger.info("Запуск записи сводных данных...")
    all_workers_data = []
    queryset = BuildObject.objects.all()
    if not queryset.exists():
        logger.warning("Нет объектов BuildObject для обработки")
        return
    for obj in queryset:
        try:
            # 1. KPI по пользователям
            table = Sheet1Table(obj)
            workers_data = table.write_summary_workers()
            logger.info(f"Сводка результатов работников записана для '{obj.name}'")
            # print(f"workers_data -> {workers_data}")
            all_workers_data.extend(workers_data)
            sleep(5)
            # 2. Сводка по объекту
            table.write_summary()
            logger.info(f"Сводная таблица за месяц записана для '{obj.name}'")
            sleep(5)

        except Exception as e:
            logger.error(f"Ошибка при обработке '{obj.name}': {e}")

    # Формируем данные для записи, преобразуем список в словарь
    data_dict = defaultdict(float)
    for row in all_workers_data:
        name, earnings = row[0], row[1]
        data_dict[name] += earnings
    if not data_dict:
        logger.warning("Нет данных для сохранения KPI")
        return
    try:
        save_to_db_current_month_kpi(data_dict)
        write_kpi_data_to_master_table(data_dict)
    except Exception as e:
        logger.error(f"Ошибка при сохранении KPI: {e}")
        raise
    logger.info("Задача завершена. Данные сохранены в БД и Google Sheets")


def save_to_db_current_month_kpi(data_dict):
    """
    Сохраняет или обновляет сводные KPI-данные по работникам за текущий месяц.

    Функция принимает список списков с данными работников, где каждая строка содержит
    имя работника и его заработок. Если работник присутствует в списке несколько раз
    (например, работал на нескольких объектах), его заработок суммируется.

    Затем функция сохраняет агрегированные данные в базу данных в виде JSON-словаря
    формата {имя: суммарный заработок} в модели MonthlyKPIData. Запись связывается с меткой
    месяца в формате 'месяц_год' (например, 'июнь_2025'). Если запись с такой меткой уже
    существует, данные в ней обновляются.

    Аргументы:
        data_dict (dict): key: имя, value: заработок.

    Возвращает:
        MonthlyKPIData: Объект модели, содержащий сохранённые данные.
    """

    # Генерируем label, например: 'июнь_2025'
    now = datetime.now()
    label = f"{now.strftime('%B').lower()}_{now.year}"

    # Ищем или создаем запись
    obj, created = MonthlyKPIData.objects.get_or_create(label=label)
    obj.data = data_dict
    obj.save()
    return obj


def write_kpi_data_to_master_table(data_dict):
    """
    Записывает агрегированные KPI-данные работников в мастер-таблицу Google Sheets.
    Функция выполняет следующие действия:
    1. Получает объект ForemanAndWorkersKPISheet за текущий год из БД
    2. Инициализирует WorkersKPIMasterTable с полученным объектом
    3. Записывает переданные KPI-данные в соответствующую таблицу Google Sheets
    Параметры:
        data_dict (dict): Словарь с данными KPI в формате:
                         {
                             "Имя работника1": суммарный_заработок,
                             "Имя работника2": суммарный_заработок,
                             ...
                         }
    Исключения:
        ForemanAndWorkersKPISheet.DoesNotExist: Если в базе нет записи за текущий год
        Exception: Любые другие ошибки при работе с Google Sheets API
    Пример использования:
        data = {"Иванов И.": 1500.50, "Петров А.": 2000.00}
        write_kpi_data_to_master_table(data)
    """
    current_year = datetime.now().year
    try:
        # Получаем объект ForemanAndWorkersKPISheet за текущий год
        kpi_sheet_obj = ForemanAndWorkersKPISheet.objects.get(year=current_year)
        workers_kpi_master_table = WorkersKPIMasterTable(kpi_sheet_obj)
        workers_kpi_master_table.write_monthly_kpi(data_dict)
    except ForemanAndWorkersKPISheet.DoesNotExist:
        error_msg = f"Не найден объект ForemanAndWorkersKPISheet за {current_year} год."
        logger.error(error_msg)
        raise ForemanAndWorkersKPISheet.DoesNotExist(error_msg)
    except Exception as e:
        logger.error(f"Ошибка при записи KPI данных в мастер-таблицу: {str(e)}")
        raise