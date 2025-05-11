from datetime import datetime
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
