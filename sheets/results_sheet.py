from datetime import datetime, date
from calendar import monthrange
from typing import List
import logging

from sheets.base_sheet import BaseTable
from buildings.models import WorkEntry, Material

logger = logging.getLogger(__name__)


class ResultsTable(BaseTable):
    """
    Класс для создания и управления таблицами результатов работы в Google Sheets.
    Создает отчеты по работникам, использованным материалам и заработку за месяц.
    """

    def __init__(self, obj):
        self.build_object = obj
        self.sheet_name = self.generate_sheet_name()
        self.sh_url = self.build_object.sh_url
        super().__init__(obj, sheet_name=self.sheet_name)

    def generate_sheet_name(self) -> str:
        """
        Генерирует название листа в формате results_<месяц>_<год>.

        Returns:
            str: Название листа
        """
        now = date.today()
        return f"results_{now.strftime('%B')}_{now.year}"

    def get_worker_monthly_data(self) -> dict:
        """
        Получает данные о работе сотрудников за текущий месяц.

        Returns:
            dict: Структура данных в формате {имя_работника: {день: [(материал, количество), ...], ...}, ...}
        """
        today = date.today()
        try:
            entries = WorkEntry.objects.filter(
                date__year=today.year,
                date__month=today.month,
                build_object=self.build_object
            ).select_related("worker")

            worker_data = {}

            for entry in entries:
                worker_name = entry.worker.name
                day = entry.date.day
                materials_dict = entry.get_materials_used()

                if worker_name not in worker_data:
                    worker_data[worker_name] = {}

                if day not in worker_data[worker_name]:
                    worker_data[worker_name][day] = []

                for material, qty in materials_dict.items():
                    worker_data[worker_name][day].append((material, qty))

            print(f"Функция получения данных о отчетах сотрудников get_worker_monthly_data вернула {worker_data}")
            return worker_data
        except Exception as e:
            logger.error(f"Ошибка при получении данных о работниках: {e}")
            return {}

    def get_materials_from_object(self) -> List[str]:
        """
        Получает список материалов, связанных с объектом строительства.

        Returns:
            List[str]: Список названий материалов
        """
        try:
            return list(self.build_object.material.values_list("name", flat=True))
        except Exception as e:
            logger.error(f"Ошибка при получении списка материалов: {e}")
            return []

    def get_material_price(self, material_name: str) -> float:
        """
        Возвращает цену материала по его названию.
        Если материал не найден или цена отсутствует - возвращает 0.

        Args:
            material_name (str): Название материала

        Returns:
            float: Цена материала
        """
        try:
            material = Material.objects.filter(name=material_name).first()
            if material and material.price is not None:
                return material.price
            return 0
        except Exception as e:
            logger.error(f"Ошибка при получении цены материала {material_name}: {e}")
            return 0

    def build_worker_table(self) -> List[List[str]]:
        """
        Создаёт таблицу отчётов для всех работников: заголовок, материалы, дни месяца и итоговые строки.
        Возвращает объединённый список всех таблиц для работников с разделителями.
        """
        materials = self.get_materials_from_object()
        workers_data = self.get_worker_monthly_data()

        all_worker_data = []

        try:
            today = date.today()
            month_name = today.strftime('%B')
            days_in_month = monthrange(today.year, today.month)[1]

            for worker_name, worker_days in workers_data.items():
                # Вычисляем, на какой строке начнётся текущая таблица
                start_row_idx = len(all_worker_data) + 1  # Sheets начинаются с 1

                values = [
                    ["Name", worker_name],
                    ["Month", month_name],
                    ["Price"] + [self.get_material_price(material) for material in materials],
                    ["Material"] + materials,
                ]

                days_row = ["Day"] + [" "] * len(materials)
                values.append(days_row)

                for day in range(1, days_in_month + 1):
                    row = [day] + [" "] * len(materials)
                    if day in worker_days:
                        for idx, (material, qty) in enumerate(worker_days[day]):
                            material_idx = materials.index(material)
                            row[material_idx + 1] = str(qty)
                    values.append(row)

                total_materials_row = ["Total materials"]
                earned_row = ["Earned"]

                # Расчёт строк и формул
                day_start_row = start_row_idx + 5
                day_end_row = day_start_row + days_in_month - 1
                price_row = start_row_idx + 2
                total_row = day_end_row + 1
                earned_total_row = total_row + 1

                for col_idx in range(len(materials)):
                    col_letter = chr(66 + col_idx)  # B, C, D...

                    total_formula = f"=SUM({col_letter}{day_start_row}:{col_letter}{day_end_row})"
                    total_materials_row.append(total_formula)

                    price_cell = f"{col_letter}{price_row}"
                    total_cell = f"{col_letter}{total_row}"
                    earned_formula = f"={total_cell}*{price_cell}"
                    earned_row.append(earned_formula)

                total_earned_formula = f"=SUM(B{earned_total_row}:{chr(65 + len(materials))}{earned_total_row})"
                total_earned_row = ["Total earned", total_earned_formula]

                values.append(total_materials_row)
                values.append(earned_row)
                values.append(total_earned_row)

                all_worker_data.extend(values)
                all_worker_data.extend([[' '], [' '], [' ']])  # Разделитель

            logger.info("Таблицы для всех работников успешно созданы")
            print(f"Функция создания таблиц отчётов для всех работников вернула {all_worker_data}\n")
            return all_worker_data

        except Exception as e:
            logger.error(f"Ошибка при создании таблиц для работников: {e}")
            return []


    # def build_worker_table(self) -> List[List[str]]:
    #     """
    #     Создаёт таблицу отчётов для всех работников: заголовок, материалы, дни месяца и итоговые строки.
    #     Возвращает объединённый список всех таблиц для работников с разделителями.
    #     """
    #     # Получаем список материалов
    #     materials = self.get_materials_from_object()
    #     # Получаем данные о работниках
    #     workers_data = self.get_worker_monthly_data()

    #     all_worker_data = []  # Список для хранения данных для всех работников

    #     try:
    #         # Получаем текущий месяц и год
    #         today = date.today()
    #         month_name = today.strftime('%B')
    #         days_in_month = monthrange(today.year, today.month)[1]

    #         # Для каждого работника создаём таблицу
    #         for worker_name, worker_days in workers_data.items():
    #             values = [
    #                 ["Name", worker_name],
    #                 ["Month", month_name],
    #                 ["Price"] + [self.get_material_price(material) for material in materials],
    #                 ["Material"] + materials,
    #             ]

    #             # Строка с днями месяца — пустые ячейки для каждого материала
    #             days_row = ["Day"] + [" "] * len(materials)
    #             values.append(days_row)

    #             # Пустые строки для каждого дня
    #             for day in range(1, days_in_month + 1):
    #                 row = [day] + [" "] * len(materials)
    #                 # Заполняем ячейки материалами, если есть данные для этого дня
    #                 if day in worker_days:
    #                     for idx, (material, qty) in enumerate(worker_days[day]):
    #                         material_idx = materials.index(material)  # Индекс материала в списке
    #                         row[material_idx + 1] = str(qty)  # Вставляем количество материала в соответствующую ячейку
    #                 values.append(row)

    #             # Строки итогов
    #             total_materials_row = ["Total materials"]
    #             earned_row = ["Earned"]

    #             for col_idx in range(len(materials)):
    #                 col_letter = chr(66 + col_idx)  # B, C, D, ...
    #                 total_formula = f"=SUM({col_letter}6:{col_letter}{5 + days_in_month})"
    #                 total_materials_row.append(total_formula)

    #                 price_cell = f"{col_letter}3"
    #                 total_cell = f"{col_letter}{6 + days_in_month}"
    #                 earned_formula = f"={total_cell}*{price_cell}"
    #                 earned_row.append(earned_formula)

    #             total_earned_formula = f"=SUM(B{7 + days_in_month}:{chr(65 + len(materials))}{7 + days_in_month})"
    #             total_earned_row = ["Total earned", total_earned_formula]

    #             values.append(total_materials_row)
    #             values.append(earned_row)
    #             values.append(total_earned_row)

    #             # Добавляем данные для работника в общий список
    #             all_worker_data.extend(values)
    #             # Добавляем три пустые строки между отчётами
    #             all_worker_data.extend([[' '], [' '], [' ']])

    #         logger.info("Таблицы для всех работников успешно созданы")
    #         print(f"Функция создания таблиц отчётов для всех работников вернула {all_worker_data}\n")
    #         return all_worker_data

    #     except Exception as e:
    #         logger.error(f"Ошибка при создании таблиц для работников: {e}")
    #         return []

    def write_to_google_sheets(self, all_worker_data: List[List[str]]):
        """
        Записывает данные о работниках и их материалах в целевой лист Google Sheets.

        Args:
            all_worker_data (List[List[str]]): Данные для записи в таблицу
        """
        try:
            # Получаем ID таблицы с помощью метода get_spreadsheet_id
            spreadsheet_id = self.get_spreadsheet_id()

            # Запись данных в таблицу
            range_ = f"{self.sheet_name}!A1"  # Начинаем запись с ячейки A1
            body = {
                "values": all_worker_data
            }

            # Отправляем запрос на запись данных
            response = self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_,
                valueInputOption="USER_ENTERED",  # Указываем формат записи
                body=body
            ).execute()

            logger.info(f"Данные успешно записаны в Google Sheets: {response}")
            print(f"Данные успешно записаны в Google Sheets: {response}")

            format_ranges = self.get_format_ranges(all_worker_data)

            for start_row, end_row, num_columns in format_ranges:
                self.apply_styles(
                    start_row=start_row,
                    end_row=end_row,
                    start_col=0,  # например, начинать с колонки B (1 = B)
                    end_col=num_columns
                )

        except Exception as e:
            logger.error(f"Произошла ошибка при записи в Google Sheets: {e}")
            print(f"Произошла ошибка при записи в Google Sheets: {e}")

    # @staticmethod
    # def get_format_ranges(data):
    #     print(f"Функция get_format_ranges получила данные {data}")
    #     ranges = []
    #     start_row = None
    #     num_columns = 0
    #
    #     for idx, row in enumerate(data):
    #         first_cell = row[0] if row else ''
    #         if first_cell and start_row is None:
    #             start_row = idx  # нашли начало блока
    #             if first_cell == 'Material':
    #                 num_columns = len(row) - 1  # минус 1, т.к. 'Material' — это заголовок
    #         elif not first_cell and start_row is not None:
    #             end_row = idx - 1
    #             ranges.append((start_row, end_row, num_columns))
    #             start_row = None
    #
    #     if start_row is not None:
    #         end_row = len(data) - 1
    #         ranges.append((start_row, end_row, num_columns))
    #     print(f"Функция get_format_ranges вернула границы {ranges}")
    #     return ranges


    @staticmethod
    def get_format_ranges(data):
        format_ranges = []
        total_rows = len(data)
        row_idx = 1

        while row_idx < total_rows:
            # Ищем начало блока
            while row_idx < total_rows and not (data[row_idx] and str(data[row_idx][0]).strip()):
                row_idx += 1
            if row_idx >= total_rows:
                break
            start_row = row_idx

            # Ищем строку "Material" для определения количества колонок
            material_columns_count = next(
                (len(data[temp_idx]) for temp_idx in range(start_row, min(start_row + 10, total_rows))
                 if data[temp_idx] and str(data[temp_idx][0]).strip().lower() == 'material'),
                None
            )

            # Ищем конец блока
            row_idx += 1
            while row_idx < total_rows and (data[row_idx] and str(data[row_idx][0]).strip()):
                row_idx += 1
            end_row = row_idx

            if material_columns_count is not None:
                format_ranges.append((start_row, end_row, material_columns_count))

            # Пропускаем пустые строки между блоками
            while row_idx < total_rows and not (data[row_idx] and str(data[row_idx][0]).strip()):
                row_idx += 1

        # Корректируем стартовые индексы для второго и последующих блоков
        corrected_ranges = [
            (start_row + 1, end_row, material_columns_count) if idx > 0 else (start_row, end_row,
                                                                              material_columns_count)
            for idx, (start_row, end_row, material_columns_count) in enumerate(format_ranges)
        ]
        print(f"Функция get_format_ranges вернула {corrected_ranges}")
        return corrected_ranges
