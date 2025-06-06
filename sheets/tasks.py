import json
import os
from datetime import datetime, date, timedelta
from pytz import timezone
from time import sleep

import pytz
import logging
from celery import shared_task

from sheets.object_kpi_sheet import ObjectKPITable
from sheets.photos_sheet import PhotosTable
from sheets.results_sheet import ResultsTable
from sheets.sheet1 import Sheet1Table
from sheets.users_kpi_sheet import UsersKPITable
from sheets.work_time_sheet import WorkTimeTable
from tsa_app.models import BuildObject

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
    # --- Запуск только в конце месяца (временно отключено) ---
    # edmonton_tz = timezone("America/Edmonton")
    # now_in_edmonton = datetime.now(edmonton_tz)
    #
    # if not is_last_day_of_month(now_in_edmonton):
    #     logger.info("Сегодня не последний день месяца по времени Калгари. Задача завершена без выполнения.")
    #     return
    # ---------------------------------------------------------
    logger.info("Сегодня последний день месяца по времени Калгари. Запуск записи сводных данных...")

    queryset = BuildObject.objects.all()
    for obj in queryset:
        try:
            # 1. KPI по пользователям
            table = Sheet1Table(obj)
            table.write_summary_workers()
            logger.info(f"Сводка результатов работников записана для '{obj.name}'")
            sleep(10)
            # 2. Сводка по объекту
            table.write_summary()
            logger.info(f"Сводная таблица за месяц записана для '{obj.name}'")
            sleep(10)

        except Exception as e:
            logger.error(f"Ошибка при обработке '{obj.name}': {e}")

    logger.info("Месячная задача завершена.")