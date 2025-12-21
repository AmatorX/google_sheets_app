import json
import os
from calendar import monthrange
from collections import defaultdict
from datetime import datetime, date, timedelta
from time import sleep
import logging
from celery import shared_task

from sheets.all_workers_kpi_shhet import WorkersKPIMasterTable
from sheets.general_statistics_sheet import GeneralStatisticTable
from sheets.object_kpi_sheet import ObjectKPITable
from sheets.photos_sheet import PhotosTable
from sheets.results_sheet import ResultsTable
from sheets.sheet1 import Sheet1Table
from sheets.users_kpi_sheet import UsersKPITable
from sheets.work_time_sheet import WorkTimeTable
from tsa_app.models import BuildObject, MonthlyKPIData, ForemanAndWorkersKPISheet, WorkEntry, Material, \
    GeneralStatisticSheet, WorkSpecialization

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
    build_objects = BuildObject.objects.filter(is_archived=False)

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
            logger.error(f"Error processing BuildObject {obj.id} ({obj.name}): {e}", exc_info=True)
            continue

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(result_data, f, ensure_ascii=False, indent=2)

    return f"Готово. Данные сохранены в {filepath}"


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
    queryset = BuildObject.objects.all()
    if not queryset.exists():
        logger.warning("Нет объектов BuildObject для обработки")
        return
    for obj in queryset:
        try:
            # 1. KPI по пользователям
            table = Sheet1Table(obj)
            table.write_summary_workers()
            logger.info(f"Сводка результатов работников записана для '{obj.name}'")
            sleep(5)
            # 2. Сводка по объекту
            table.write_summary()
            logger.info(f"Сводная таблица за месяц записана для '{obj.name}'")
            sleep(5)

        except Exception as e:
            logger.error(f"Ошибка при обработке '{obj.name}': {e}")

    # Формируем данные для записи
    data_dict = get_worker_materials_hourly_earnings()
    try:
        save_to_db_current_month_kpi(data_dict)
        write_kpi_data_to_master_table(data_dict)
    except Exception as e:
        logger.error(f"Ошибка при сохранении KPI: {e}")
        raise
    logger.info("Задача завершена. Данные сохранены в БД и Google Sheets")


############################################################################################################
@shared_task
def update_foremans_and_workers_kpi():
    for month in range(1, 13):
        # Формируем данные для записи
        data_dict = get_worker_materials_hourly_earnings(month)
        try:
            # save_to_db_current_month_kpi(data_dict)
            write_kpi_data_to_master_table(data_dict, month=month)
        except Exception as e:
            logger.error(f"Ошибка при сохранении KPI: {e}")
            raise
        logger.info("Задача завершена. Данные сохранены в БД и Google Sheets")
############################################################################################################

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


def write_kpi_data_to_master_table(data_dict: dict, month=None):
    """
    Записывает агрегированные KPI-данные работников в мастер-таблицу Google Sheets.

    Параметры:
        data_dict (dict): Словарь с данными KPI в формате:
                         {
                             "Имя работника1": суммарный_заработок,
                             "Имя работника2": суммарный_заработок,
                             ...
                         }
        month (int | None): Номер месяца (1–12). Если None — используется текущий месяц.

    Исключения:
        ForemanAndWorkersKPISheet.DoesNotExist: Если в базе нет записи за текущий год
        Exception: Любые другие ошибки при работе с Google Sheets API
    """
    current_year = datetime.now().year
    try:
        # Получаем объект ForemanAndWorkersKPISheet за текущий год
        kpi_sheet_obj = ForemanAndWorkersKPISheet.objects.get(year=current_year)

        workers_kpi_master_table = WorkersKPIMasterTable(kpi_sheet_obj)
        if month == 1:
            workers_kpi_master_table.clear_sheet()
        workers_kpi_master_table.write_monthly_kpi(data_dict, month=month)

    except ForemanAndWorkersKPISheet.DoesNotExist:
        error_msg = f"Не найден объект ForemanAndWorkersKPISheet за {current_year} год."
        logger.error(error_msg)
        raise ForemanAndWorkersKPISheet.DoesNotExist(error_msg)

    except Exception as e:
        logger.error(f"Ошибка при записи KPI данных в мастер-таблицу: {str(e)}")
        raise


def get_worker_materials_hourly_earnings(month: int | None = None, year: int | None = None) -> dict:
    """
    Возвращает словарь {worker_name: заработок в час за установленные материалы}
    за указанный месяц и год.

    Если month или year не указаны — используется текущий месяц и год.
    """

    today = date.today()
    month = month or today.month
    year = year or today.year

    start_date = date(year, month, 1)
    end_date = date(year, month, monthrange(year, month)[1])

    # Получаем рабочие записи за период только для неархивных работников
    entries = (
        WorkEntry.objects
        .select_related('worker')
        .filter(
            date__range=(start_date, end_date),
            worker__is_archived=False,
        )
    )

    # Цены материалов {material_name: price}
    material_prices = {
        material.name: material.price
        for material in Material.objects.filter(price__isnull=False)
    }

    worker_totals = defaultdict(lambda: {'total_price': 0.0, 'total_hours': 0.0})

    for entry in entries:
        worker_name = entry.worker.name
        worked_hours = entry.get_worked_hours_as_float()
        materials_used = entry.get_materials_used()

        materials_cost = sum(
            float(material_prices.get(material_name, 0)) * float(quantity)
            for material_name, quantity in materials_used.items()
        )
        worker_totals[worker_name]['total_price'] += materials_cost
        worker_totals[worker_name]['total_hours'] += worked_hours

    # Итог: заработок в час
    return {
        worker_name: round(
            data['total_price'] / data['total_hours'], 2
        ) if data['total_hours'] else 0.0
        for worker_name, data in worker_totals.items()
    }


@shared_task
def update_general_statistic_task():
    try:
        summary = GeneralStatisticSheet.objects.get()
    except GeneralStatisticSheet.DoesNotExist:
        logger.error("Не найден объект GeneralStatistic. Создай его через админку.")
        return
    work_specializations = WorkSpecialization.objects.all()
    for work_spec in work_specializations:
        table = GeneralStatisticTable(summary, work_spec)
        table.ensure_sheet_exists()
        table.write_full_statistic()
    table = GeneralStatisticTable(summary)
    table.ensure_sheet_exists()
    table.write_full_statistic()