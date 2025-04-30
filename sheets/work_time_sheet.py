import logging
import calendar

from buildings.models import WorkEntry
from collections import defaultdict

from utils.chunk import create_chunk_data
from sheets.base_sheet import BaseTable

logger = logging.getLogger(__name__)


class WorkTimeTable(BaseTable):
    def __init__(self, build_object, sheet_name='Work Time'):
        super().__init__(build_object, sheet_name)

    def create_table(self, date_chunk, rows):
        """
        Создаёт таблицу рабочих часов на основе переданного чанка и данных.
        """
        self.ensure_sheet_exists()

        start_row = self.get_last_non_empty_row() + 3  # отступ между таблицами
        logger.info(f"Старт записи таблицы 'Work Time' со строки {start_row}")

        num_days = len(date_chunk)
        headers = ["Per hour", "User Name"] + [f"{month} {day}" for month, _, day in date_chunk] + ["Total", "Salary"]
        values = [headers]

        row_index = start_row + 1
        for user in rows:
            work_hours = user["work_hours"]
            total_formula = f"=SUM(C{row_index}:P{row_index})"
            salary_formula = f"=Q{row_index}*A{row_index}"
            values.append([user["per_hour"], user["user_name"], *work_hours, total_formula, salary_formula])
            row_index += 1

        # Итого
        total_sum_formula = f"=SUM(Q{start_row + 1}:Q{row_index - 1})"
        salary_sum_formula = f"=SUM(R{start_row + 1}:R{row_index - 1})"
        values.append(["", "Total"] + [""] * num_days + [total_sum_formula, salary_sum_formula])

        # Запись
        self.update_data(f"A{start_row}", values)
        logger.info(f"Таблица 'Work Time' обновлена: {len(rows)} работников, строки {start_row}-{row_index}")

        # Стили
        self.apply_styles(start_row, row_index, 0, len(headers))
        print(f'Таблица для объекта {self.build_object.name} создана\n')

    def remove_existing_chunk_if_exists(self, chunk):
        sheet_name = self.sheet_name
        range_name = f"{sheet_name}!C1:C"
        first_day_str = f"{chunk[0][0]} {chunk[0][2]}"  # например, "April 6"

        # Получаем значения из колонки C
        response = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=range_name
        ).execute()

        col_values = response.get('values', [])

        # Поиск строки с первым днём чанка
        start_row = None
        for idx, row in enumerate(col_values):
            if row and row[0].strip() == first_day_str:
                print(f'Начало текущего чанка в строке {idx + 1}')
                start_row = idx
                break

        if start_row is None:
            print("Текущий чанк не найден — ничего не удаляем.")
            return

        # Определяем границу таблицы — пока есть непустые строки
        end_row = start_row
        for i in range(start_row + 1, len(col_values)):
            if col_values[i] and col_values[i][0].strip():
                end_row = i
            else:
                break

        # Удалим строки от (start_row) до (end_row + 2)
        delete_start = max(0, start_row )
        delete_end = end_row + 2

        print(f"Удаляем строки с {delete_start + 1} по {delete_end + 1}")

        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={
                "requests": [
                    {
                        "deleteDimension": {
                            "range": {
                                "sheetId": self.sheet_id,
                                "dimension": "ROWS",
                                "startIndex": delete_start,
                                "endIndex": delete_end + 1  # endIndex не включительно
                            }
                        }
                    }
                ]
            }
        ).execute()
        print("Удаление завершено.")

    def build_worktime_rows(self, chunk_data):
        """
        Генерирует rows (список словарей) для таблицы рабочих часов
        на основе WorkEntry и данных чанка, ограничиваясь текущим build_object.
        """
        from datetime import datetime
        from collections import defaultdict
        import calendar

        # Преобразуем chunk_data в список дат
        year = datetime.now().year
        dates = []
        for month_name, weekday, day_str in chunk_data:
            day = int(day_str)
            month = list(calendar.month_name).index(month_name)
            date_obj = datetime(year, month, day).date()
            dates.append(date_obj)

        # Получаем WorkEntry только для текущего build_object и нужных дат
        work_entries = WorkEntry.objects.filter(
            build_object=self.build_object,
            date__in=dates
        ).select_related("worker")

        # Структура: {worker_id: {date: hours}}
        worker_data = defaultdict(lambda: defaultdict(float))
        per_hour_map = {}

        for entry in work_entries:
            worker = entry.worker
            worker_data[worker.id][entry.date] += float(entry.worked_hours)
            per_hour_map[worker.id] = worker.salary

        # Формируем rows
        rows = []
        for worker_id, date_hours in worker_data.items():
            user_name = WorkEntry.objects.filter(worker_id=worker_id).first().worker.name
            work_hours = [date_hours.get(date, 0) for date in dates]
            rows.append({
                "per_hour": per_hour_map[worker_id],
                "user_name": user_name,
                "work_hours": work_hours
            })

        return rows

    def generate_current_chunk_data_and_rows(self):
        """
        Генерирует чанк и данные для текущего 14-дневного периода.
        Использует self.build_worktime_rows для получения данных.
        """
        chunk = create_chunk_data(full_list=False)
        rows = self.build_worktime_rows(chunk)
        print(f'generate_current_chunk_data_and_rows вернул \n'
              f'chunk: {chunk}\n'
              f'rows: {rows}')
        return chunk, rows

    # @staticmethod
    # def generate_current_chunk_data_and_rows():
    #     """
    #     Генерирует чанк и данные для текущего 14-дневного периода.
    #     Может использоваться извне для подготовки данных.
    #     """
    #     chunk = create_chunk_data(full_list=False)
    #     rows = build_worktime_rows(chunk)
    #     print(f'generate_current_chunk_data_and_rows вернул \n'
    #           f'chunk: {chunk}\n'
    #           f'rows: {rows}')
    #     return chunk, rows



# def build_worktime_rows(chunk_data):
#     """
#     Генерирует rows (список словарей) для таблицы рабочих часов на основе WorkEntry и данных чанка.
#     """
#     from datetime import datetime
#
#     # Преобразуем chunk_data в список дат
#     year = datetime.now().year
#     dates = []
#     for month_name, weekday, day_str in chunk_data:
#         day = int(day_str)
#         month = list(calendar.month_name).index(month_name)
#         date_obj = datetime(year, month, day).date()
#         dates.append(date_obj)
#
#     # Получаем все записи WorkEntry в пределах этих дат
#     work_entries = WorkEntry.objects.filter(date__in=dates).select_related("worker")
#
#     # Структура: {worker_id: {date: hours}}
#     worker_data = defaultdict(lambda: defaultdict(float))
#     per_hour_map = {}
#
#     for entry in work_entries:
#         worker = entry.worker
#         worker_data[worker.id][entry.date] += float(entry.worked_hours)
#         per_hour_map[worker.id] = worker.salary  # или другой способ получить ставку
#
#     # Формируем rows
#     rows = []
#     for worker_id, date_hours in worker_data.items():
#         user_name = WorkEntry.objects.filter(worker_id=worker_id).first().worker.name
#         work_hours = [date_hours.get(date, 0) for date in dates]
#         rows.append({
#             "per_hour": per_hour_map[worker_id],
#             "user_name": user_name,
#             "work_hours": work_hours
#         })
#
#     return rows
