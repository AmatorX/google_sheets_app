import logging
from sheets.base_sheet import BaseTable

logger = logging.getLogger(__name__)


class WorkTimeTable(BaseTable):
    def __init__(self, spreadsheet_url, sheet_name='Work Time'):
        super().__init__(spreadsheet_url, sheet_name)

    def create_table(self, date_chunk, rows):
        """Создаёт таблицу рабочих часов"""
        self.ensure_sheet_exists()
        start_row = self.get_last_non_empty_row() + 3  # + 3 чтобы отступить строку между таблицами
        print(f'start_row: {start_row}')
        logger.info(f'Старт записи таблицы со строки start_row: {start_row}')
        num_days = len(date_chunk)
        headers = ["Per hour", "User Name"] + [f"{month} {day}" for month, _, day in date_chunk] + ["Total", "Salary"]
        values = [headers]

        row_index = start_row + 1  # Начало данных после заголовка
        for user in rows:
            work_hours = user["work_hours"]
            total_formula = f"=SUM(C{row_index}:P{row_index})"
            # total_formula = f"=SUM(C{row_index}:Q{row_index + num_days - 1})"
            salary_formula = f"=Q{row_index}*A{row_index}"
            values.append([user["per_hour"], user["user_name"], *work_hours, total_formula, salary_formula])
            row_index += 1

        # Итоговые суммы
        total_sum_formula = f"=SUM(Q{start_row + 1}:Q{row_index - 1})"
        salary_sum_formula = f"=SUM(R{start_row + 1}:R{row_index - 1})"
        values.append(["", "Total"] + [""] * num_days + [total_sum_formula, salary_sum_formula])

        # Запись данных
        self.update_data(f"A{start_row}", values)
        logger.info(f"Таблица 'Work Time' обновлена с {len(rows)} работниками, начиная со строки {start_row}.")

        # Применение стилей
        color = {'red': 0.85, 'green': 0.93, 'blue': 0.98}
        print(f'Начало форматирвания: {start_row}, конец форматирования: {row_index}')
        self.apply_styles(start_row, row_index, 0, len(headers), color)

