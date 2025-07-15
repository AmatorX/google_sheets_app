import logging
from collections import defaultdict
from datetime import date
from time import sleep

from sheets.base_sheet import BaseTable
from sheets.sheet1 import Sheet1Table
from tsa_app.models import BuildObject

logger = logging.getLogger(__name__)


class GeneralStatisticTable(BaseTable):
    def __init__(self, obj):
        super().__init__(obj=obj)

    def get_summary_data_for_object(self, obj, year=None):
        """
        Возвращает словарь с данными по каждому месяцу для указанного объекта.
        Формат:
        {
            'January': [[...], ...],
            'February': [[...], ...],
            ...
        }
        Строка с названием месяца в начале мини-блока удаляется.
        """
        if year is None:
            year = date.today().year

        month_names = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ]

        summary_data = {}
        table = Sheet1Table(obj)

        for month_number in range(1, 13):
            rows = table.get_summary_rows(month=month_number, year=year)
            if rows:
                summary_data[month_names[month_number - 1]] = rows[1:]

        return summary_data

    def generate_horizontal_block_for_object(self, months_data: dict[str, list[list]]) -> list[list]:
        """
        Формирует горизонтальный блок из 12 мини-таблиц по месяцам.
        Каждый месяц = 2 колонки + 1 пустая. Если данных нет — вставляются пустые ячейки.
        """
        month_names = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ]

        header_row = []
        data_rows = [[] for _ in range(5)]

        for month in month_names:
            if month in months_data:
                mini_block = months_data[month]
                header_row.extend([month, '', ''])
                for i in range(5):
                    row = mini_block[i] if i < len(mini_block) else ['', '']
                    data_rows[i].extend(row + [''])
            else:
                header_row.extend(['', '', ''])
                for i in range(5):
                    data_rows[i].extend(['', '', ''])

        return [header_row] + data_rows

    def write_objects_statistic_table(self, start_row: int) -> dict:
        """
        Записывает блоки с мини-таблицами по каждому объекту.
        Также собирает агрегированные данные по Salary, Earned, Result для всех месяцев.
        Возвращает словарь:
        {
            'Salary': {month: value, ...},
            'Earned': {...},
            'Result': {...}
        }
        """
        build_objects = BuildObject.objects.all()
        all_rows = []
        style_blocks = []
        current_row = start_row

        months = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]

        summary_totals = {
            'Salary': defaultdict(float),
            'Earned': defaultdict(float),
            'Result': defaultdict(float),
        }

        for obj in build_objects:
            logger.info(f"Обработка объекта: {obj.name}")
            obj_data = self.get_summary_data_for_object(obj)
            block = self.generate_horizontal_block_for_object(obj_data)
            block_width = len(block[0])

            for month in months:
                mini_block = obj_data.get(month)
                if not mini_block or len(mini_block) < 4:
                    continue
                try:
                    summary_totals['Salary'][month] += float(mini_block[1][1])
                    summary_totals['Earned'][month] += float(mini_block[2][1])
                    summary_totals['Result'][month] += float(mini_block[3][1])
                except (ValueError, IndexError) as e:
                    logger.warning(f"Ошибка при парсинге значений для {obj.name}, {month}: {e}")

            block_header = [f'Project: "{obj.name}"'] + [''] * (block_width - 1)
            all_rows.append(block_header)
            block_start_row = current_row
            current_row += 1

            all_rows.append([''] * block_width)
            current_row += 1

            all_rows += block
            current_row += len(block)

            block_end_row = current_row
            all_rows += [[''] * block_width for _ in range(6)]
            current_row += 6

            style_blocks.append((block_start_row, block_end_row))

        self.update_data(f"A{start_row}", all_rows)

        for start, end in style_blocks:
            self.apply_styles(
                start_row=start,
                end_row=end,
                start_col=0,
                end_col=len(all_rows[0]),
                bold=True,
                borders=False,
            )
            sleep(1.1)

        return summary_totals

    def write_summary_table(self, start_row: int, summary_totals: dict) -> int:
        """
        Записывает в лист сводную таблицу со значениями Salary, Earned, Result по всем объектам и месяцам.
        Возвращает следующую строку для продолжения вставки.
        """
        months = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]

        header = ["All teams", "Total"] + months

        salary_values = [summary_totals['Salary'].get(m, 0) for m in months]
        earned_values = [summary_totals['Earned'].get(m, 0) for m in months]
        result_values = [summary_totals['Result'].get(m, 0) for m in months]

        salary_row = ["Salary", sum(salary_values)] + salary_values
        earned_row = ["Earned", sum(earned_values)] + earned_values
        result_row = ["Result", sum(result_values)] + result_values

        empty_rows = [[''] * len(header) for _ in range(2)]
        rows = [header, salary_row, earned_row, result_row] + empty_rows

        self.update_data(f"A{start_row}", rows)

        self.apply_styles(
            start_row=start_row,
            end_row=start_row + 3,
            start_col=0,
            end_col=len(header),
            bold=True,
            borders=True,
        )

        return start_row + len(rows)

    def write_full_statistic(self):
        """
        Полный процесс записи отчёта:
        1. Очистка листа
        2. Запись общей таблицы
        3. Запись таблиц по объектам
        """
        try:
            self.clear_sheet()
        except Exception as e:
            logger.error(f"Ошибка очистки листа: {e}")

        try:
            summary_totals = self.write_objects_statistic_table(start_row=10)
        except Exception as e:
            logger.error(f"Ошибка записи таблиц объектов по месяцам: {e}")
            summary_totals = {'Salary': {}, 'Earned': {}, 'Result': {}}

        try:
            self.write_summary_table(start_row=1, summary_totals=summary_totals)
        except Exception as e:
            logger.error(f"Ошибка записи общей таблицы результатов: {e}")

        logger.info("Данные успешно записаны и стилизованы.")
