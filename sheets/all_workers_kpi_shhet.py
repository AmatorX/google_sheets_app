import logging
from datetime import datetime
from time import sleep

from sheets.base_sheet import BaseTable
from utils.column_index_to_letter import column_index_to_letter

logger = logging.getLogger(__name__)


class WorkersKPIMasterTable(BaseTable):
    """
    Класс для работы с таблицами общего KPI для всех работников.
    """
    def __init__(self, obj):
        current_year = datetime.now().year
        sheet_name = f"Teams KPI {current_year} New"
        super().__init__(obj, sheet_name)

    # def write_monthly_kpi(self, workers_dict: dict[str, float]):
    #     self.ensure_sheet_exists()
    #
    #     month_name = datetime.now().strftime("%B")  # Например, "June"
    #     # month_name = "March"
    #
    #     # Получаем заголовки
    #     try:
    #         header_range = f"{self.sheet_name}!A1:1"
    #         result = self.service.spreadsheets().values().get(
    #             spreadsheetId=self.spreadsheet_id,
    #             range=header_range
    #         ).execute()
    #         sleep(1)
    #         headers = result.get("values", [[]])[0]
    #     except Exception as e:
    #         logger.error(f"Ошибка при получении заголовков: {e}")
    #         headers = []
    #
    #     # Если заголовков нет — создаём
    #     if not headers or headers[0] != "№":
    #         headers = ["№", "Employee's name", month_name]
    #         self.update_data("A1", [headers])
    #         existing_names = []
    #     else:
    #         # Получаем имена работников из столбца B (второй столбец)
    #         try:
    #             name_range = f"{self.sheet_name}!B2:B"
    #             result = self.service.spreadsheets().values().get(
    #                 spreadsheetId=self.spreadsheet_id,
    #                 range=name_range
    #             ).execute()
    #             sleep(1)
    #             existing_names = [row[0] for row in result.get("values", [])]
    #         except Exception as e:
    #             logger.error(f"Ошибка при получении имён работников: {e}")
    #             existing_names = []
    #
    #     # Добавляем столбец для месяца, если его нет
    #     if month_name in headers:
    #         month_col_idx = headers.index(month_name)
    #     else:
    #         month_col_idx = len(headers)
    #         col_letter = column_index_to_letter(month_col_idx + 1)
    #         self.update_data(f"{col_letter}1", [[month_name]])
    #         sleep(1)
    #         headers.append(month_name)
    #
    #     # Запись KPI
    #     name_to_row = {name: idx + 2 for idx, name in enumerate(existing_names)}  # строки начинаются с 2
    #     for name, kpi in workers_dict.items():
    #         if name in name_to_row:
    #             row = name_to_row[name]
    #         else:
    #             # Новый работник — добавляем имя и нумерацию
    #             row = len(existing_names) + 2
    #             self.update_data(f"A{row}", [[row - 1]])  # Нумерация (строка - 1)
    #             self.update_data(f"B{row}", [[name]])
    #             sleep(1)# Имя работника
    #             existing_names.append(name)
    #             name_to_row[name] = row
    #         # Пишем KPI
    #         col_letter = column_index_to_letter(month_col_idx + 1)
    #         self.update_data(f"{col_letter}{row}", [[kpi]])
    #         sleep(1)
    #
    #     # После всех записей применяем стили к диапазону KPI месяца
    #     if existing_names:
    #         first_data_row = 1
    #         last_data_row = len(existing_names)  # включительно (например, 3 работника: строки 2,3,4)
    #         col_start = 0
    #         col_end = month_col_idx + 1
    #         self.apply_styles(first_data_row, last_data_row + 1, col_start, col_end)
    #
    #     logger.info(f"KPI за {month_name} записаны для {len(workers_dict)} работников.")

    def write_monthly_kpi(self, workers_dict: dict[str, float]):
        self.ensure_sheet_exists()
        month_name = datetime.now().strftime("%B")

        # Получаем все данные листа за один запрос
        try:
            full_range = f"{self.sheet_name}!A1:Z"  # Предполагаем, что данных не больше 26 столбцов
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=full_range
            ).execute()
            sleep(1)
            sheet_data = result.get("values", [])
        except Exception as e:
            logger.error(f"Ошибка при получении данных листа: {e}")
            sheet_data = []

        # Обработка заголовков
        headers = sheet_data[0] if sheet_data else []

        # Если заголовков нет или структура неверная - инициализируем
        if not headers or headers[0] != "№":
            headers = ["№", "Employee's name", month_name]
            sheet_data = [headers]
            existing_names = []
        else:
            # Получаем существующие имена из столбца B
            existing_names = [row[1] for row in sheet_data[1:] if len(row) > 1]

        # Определяем индекс столбца для месяца
        if month_name in headers:
            month_col_idx = headers.index(month_name)
        else:
            month_col_idx = len(headers)
            headers.append(month_name)
            if len(sheet_data) == 0:
                sheet_data.append(headers)
            else:
                sheet_data[0] = headers

        # Подготавливаем данные для обновления
        updates = []
        name_to_row = {name: idx + 2 for idx, name in enumerate(existing_names)}

        # Обновляем данные в памяти
        for name, kpi in workers_dict.items():
            if name in name_to_row:
                row = name_to_row[name]
                # Убедимся, что строка достаточно длинная
                while len(sheet_data) < row:
                    sheet_data.append([])
                while len(sheet_data[row - 1]) <= month_col_idx:
                    sheet_data[row - 1].append("")
                sheet_data[row - 1][month_col_idx] = kpi
            else:
                # Добавляем нового работника
                row = len(sheet_data) + 1
                new_row = [row - 1, name] + [""] * (month_col_idx - 1)
                if month_col_idx >= len(new_row):
                    new_row.append(kpi)
                else:
                    new_row[month_col_idx] = kpi
                sheet_data.append(new_row)
                existing_names.append(name)
                name_to_row[name] = row

        # Формируем один запрос на обновление всех данных
        update_range = f"{self.sheet_name}!A1:{column_index_to_letter(len(headers))}{len(sheet_data)}"
        updates.append({
            'range': update_range,
            'values': sheet_data
        })

        # Применяем все обновления одним запросом
        if updates:
            body = {
                'valueInputOption': 'USER_ENTERED',
                'data': updates
            }
            try:
                self.service.spreadsheets().values().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body=body
                ).execute()
                sleep(1)
            except Exception as e:
                logger.error(f"Ошибка при массовом обновлении данных: {e}")

        # Применяем стили (если есть данные)
        if existing_names:
            first_data_row = 1
            last_data_row = len(existing_names)
            col_start = 0
            col_end = month_col_idx + 1
            self.apply_styles(first_data_row, last_data_row + 1, col_start, col_end, bold=True)

        logger.info(f"KPI за {month_name} записаны для {len(workers_dict)} работников.")