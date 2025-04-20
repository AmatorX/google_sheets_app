from datetime import datetime
import logging
from sheets.base_sheet import BaseTable
from buildings.models import WorkEntry


logger = logging.getLogger(__name__)


class PhotosTable(BaseTable):
    """
    Класс для работы с фотоотчётами в Google Таблице.
    Создаёт таблицы по пользователям на листе формата 'photos_<месяц>_<год>'.
    """
    
    def __init__(self, build_object):
        self.build_object = build_object
        self.sheet_name = self.generate_sheet_name()
        self.spreadsheet_url = self.build_object.sh_url
        super().__init__(build_object, sheet_name=self.sheet_name)

    def generate_sheet_name(self):
        """
        Генерирует название листа в формате photos_<Month>_<Year>.

        Returns:
            str: Название листа.
        """
        now = datetime.now()
        return f"photos_{now.strftime('%B')}_{now.year}"

    def get_workers(self):
        """
        Возвращает список имён пользователей, связанных с объектом строительства.
        """
        workers = self.get_workers_for_build_object()
        return [w.name for w in workers if w.name]

    def get_usernames_from_sheet(self):
        """
        Получает список имён пользователей из колонки B текущего листа.

        Returns:
            list[str]: Имена пользователей, уже добавленные в таблицу.
        """
        try:
            range_name = f"{self.sheet_name}!B1:B"
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()
            values = result.get('values', [])
            usernames = [row[0] for row in values if row]  # Убираем пустые строки
            return usernames
        except Exception as e:
            logger.error(f"Ошибка при получении имён пользователей из листа '{self.sheet_name}': {e}")
            return []

    def create_tables_for_missing_workers(self):

        """
        Создаёт таблицы для новых пользователей, если они не присутствуют на листе.
        Каждая таблица включает дни месяца и имя работника, а также отступ между таблицами.
        """

        linked_workers = self.get_workers()
        existing_workers = self.get_usernames_from_sheet()
        missing_workers = [w for w in linked_workers if w not in existing_workers]

        worker_tables = []  # Список таблиц по каждому работнику
        style_ranges = []

        # current_row = 1  # Индекс с 1
        current_row = self.get_last_non_empty_row(column_letter='A') + 1

        for worker in missing_workers:
            table_data, style_range = self.create_worker_data(worker, current_row)
            worker_tables.append(table_data)
            style_ranges.append(style_range)
            current_row += len(table_data)

        if worker_tables:
            self.batch_update_columns(worker_tables)

        # Цвет для таблиц
        color = {'red': 0.85, 'green': 0.93, 'blue': 0.98}
        for start_row, end_row, start_col, end_col in style_ranges:
            self.apply_styles(start_row, end_row, start_col, end_col, color=color)

    def create_worker_data(self, worker_name, start_row=None):

        """
        Генерирует таблицу для конкретного работника с заголовком и строками по дням месяца.

        Args:
            worker_name (str): Имя работника.
            start_row (int, optional): Строка, с которой начинать вставку.

        Returns:
            tuple: (table_data (list[list[str]]), style_range (tuple[int, int, int, int]))
        """

        from calendar import monthrange
        now = datetime.now()
        days_in_month = monthrange(now.year, now.month)[1]

        # Если start_row не передан, определить его автоматически
        if start_row is None:
            last_row = self.get_last_non_empty_row(column_letter='A')
            start_row = last_row + 1  # следующая строка после последней непустой

        table_data = [['Day', worker_name]]
        for day in range(1, days_in_month + 1):
            table_data.append([str(day)])

        # Стиль применяется только к "основной" таблице (без отступов)
        styled_start_row = start_row
        styled_end_row = start_row + len(table_data) - 1

        # Добавляем отступ
        table_data.extend([[' '], [' '], [' ']])

        return table_data, (styled_start_row, styled_end_row, 0, 26)

    def batch_update_columns(self, worker_tables):
        """
        Пакетно добавляет таблицы в Google Sheet, используя appendCells.

        Args:
            worker_tables (list[list[list[str]]]): Список таблиц для каждого работника.
        """
        try:
            sheet_id = self.get_sheet_id()
            if sheet_id is None:
                logger.error("Не удалось получить sheetId, запись невозможна.")
                return

            requests = []

            for table_data in worker_tables:
                rows = []
                for row in table_data:
                    # Если строка пустая (например, []), всё равно добавляем как пустую строку
                    if row:
                        values = [{"userEnteredValue": {"stringValue": str(cell)}} for cell in row]
                    else:
                        values = []  # просто пустая строка

                    rows.append({"values": values})

                requests.append({
                    "appendCells": {
                        "sheetId": sheet_id,
                        "rows": rows,
                        "fields": "userEnteredValue"
                    }
                })

            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": requests}
            ).execute()

            logger.info(f"Добавлено {len(worker_tables)} таблиц (работников).")

        except Exception as e:
            logger.error(f"Ошибка при добавлении таблиц: {e}")
