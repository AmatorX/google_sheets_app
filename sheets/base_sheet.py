import logging
from typing import List, Any
from utils.service import get_service

logger = logging.getLogger(__name__)

class BaseTable:
    """
    Базовый класс для работы с Google Sheets.

    Предоставляет функциональность для:
    - Проверки и создания листа, если он не существует.
    - Получения идентификаторов таблицы и листа.
    - Определения последней непустой строки.
    - Обновления данных в ячейках.
    - Применения стилей и границ к диапазону ячеек.

    Атрибуты:
        sh_url (str): URL Google-таблицы.
        sheet_name (str): Название листа, с которым будет происходить работа.
        service (Resource): Авторизованный Google Sheets API клиент.
        spreadsheet_id (str): ID таблицы, извлечённый из URL.
        sheet_id (int): Числовой идентификатор листа.

    Методы:
        ensure_sheet_exists(): Проверяет существование листа, создаёт при отсутствии.
        get_spreadsheet_id(): Возвращает ID таблицы из URL.
        get_sheet_id(): Получает числовой ID листа по названию.
        get_last_non_empty_row(): Возвращает индекс последней непустой строки в столбце B.
        apply_styles(start_row, end_row, start_col, end_col, color): Применяет стили и границы к диапазону ячеек.
        update_data(range_name, values): Обновляет данные в заданном диапазоне листа.
    """
    def __init__(self, obj=None, sheet_name=None):
        self.obj = obj
        self.sh_url = obj.sh_url
        self.sheet_name = sheet_name
        self.service = get_service()
        self.spreadsheet_id = self.get_spreadsheet_id()
        self.sheet_id = self.get_sheet_id()

    def ensure_sheet_exists(self):
        """
        Проверяет, существует ли лист, и создает его, если не существует.
        """
        try:
            sheet_metadata = self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
            sheet_titles = [sheet["properties"]["title"] for sheet in sheet_metadata.get("sheets", [])]

            if self.sheet_name not in sheet_titles:
                requests = [{
                    "addSheet": {
                        "properties": {
                            "title": self.sheet_name
                        }
                    }
                }]
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={"requests": requests}
                ).execute()
                logger.info(f"Лист '{self.sheet_name}' создан.")
                self.sheet_id = self.get_sheet_id()  # Обновляем sheet_id после создания листа
            else:
                logger.debug(f"Лист '{self.sheet_name}' уже существует.")
                self.sheet_id = self.get_sheet_id()  # Обновляем sheet_id в любом случае
        except Exception as e:
            logger.error(f"Ошибка при проверке или создании листа: {e}")

    def get_spreadsheet_id(self):
        """Возвращает ID таблицы из URL"""
        return self.sh_url.split("/")[5]

    def get_sheet_id(self):
        """Получает sheetId по названию листа"""
        try:
            sheet_metadata = self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
            for sheet in sheet_metadata.get('sheets', []):
                if sheet['properties']['title'] == self.sheet_name:
                    return sheet['properties']['sheetId']
            logger.warning(f"Лист '{self.sheet_name}' не найден.")
            return None
        except Exception as e:
            logger.error(f"Ошибка при получении sheetId для '{self.sheet_name}': {e}")
            return None


    def get_last_non_empty_row(self, column_letter='B'):
        """
        Возвращает индекс последней непустой строки в указанном столбце.

        Args:
            column_letter (str): Буква столбца (например, 'A', 'B').

        Returns:
            int: Индекс последней непустой строки (начиная с 1).
        """
        try:
            range_name = f"{self.sheet_name}!{column_letter}1:{column_letter}"
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()
            values = result.get('values', [])
            print(f'get_last_non_empty_row values: {len(values) if values else 0}')
            return len(values) if values else 0
        except Exception as e:
            logger.error(f"Ошибка при получении последней непустой строки в столбце {column_letter}: {e}")
            return 0


    def get_workers_for_build_object(self):
        """Возвращает всех пользователей, связанных с объектом строительства"""
        from tsa_app.models import Worker  # Импортируем модель внутри метода, чтобы избежать циклических зависимостей
        workers = Worker.objects.filter(build_obj=self.build_object)
        return workers

    # def apply_styles(self, start_row, end_row, start_col, end_col):
    #     color = {'red': 0.85, 'green': 0.93, 'blue': 0.98}
    #     requests = [
    #         {
    #             'repeatCell': {
    #                 'range': {
    #                     'sheetId': self.sheet_id,
    #                     'startRowIndex': start_row - 1,
    #                     'endRowIndex': end_row,
    #                     'startColumnIndex': start_col,
    #                     'endColumnIndex': end_col
    #                 },
    #                 'cell': {'userEnteredFormat': {'backgroundColor': color}},
    #                 'fields': 'userEnteredFormat.backgroundColor'
    #             }
    #         },
    #         {
    #             'updateBorders': {
    #                 'range': {
    #                     'sheetId': self.sheet_id,
    #                     'startRowIndex': start_row - 1,
    #                     'endRowIndex': end_row,
    #                     'startColumnIndex': start_col,
    #                     'endColumnIndex': end_col
    #                 },
    #                 'top': {'style': 'SOLID', 'width': 1},
    #                 'bottom': {'style': 'SOLID', 'width': 1},
    #                 'left': {'style': 'SOLID', 'width': 1},
    #                 'right': {'style': 'SOLID', 'width': 1},
    #                 'innerHorizontal': {'style': 'SOLID', 'width': 1},
    #                 'innerVertical': {'style': 'SOLID', 'width': 1}
    #             }
    #         }
    #     ]
    #     self.service.spreadsheets().batchUpdate(
    #         spreadsheetId=self.spreadsheet_id,
    #         body={'requests': requests}
    #     ).execute()

    def apply_styles(self, start_row, end_row, start_col, end_col, bold=False):
        color = {'red': 0.85, 'green': 0.93, 'blue': 0.98}
        requests = [
            {
                'repeatCell': {
                    'range': {
                        'sheetId': self.sheet_id,
                        'startRowIndex': start_row - 1,
                        'endRowIndex': end_row,
                        'startColumnIndex': start_col,
                        'endColumnIndex': end_col
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': color,
                            'textFormat': {
                                'bold': bold  # Добавляем жирность в зависимости от параметра
                            }
                        }
                    },
                    'fields': 'userEnteredFormat(backgroundColor,textFormat.bold)'  # Обновляем оба поля
                }
            },
            {
                'updateBorders': {
                    'range': {
                        'sheetId': self.sheet_id,
                        'startRowIndex': start_row - 1,
                        'endRowIndex': end_row,
                        'startColumnIndex': start_col,
                        'endColumnIndex': end_col
                    },
                    'top': {'style': 'SOLID', 'width': 1},
                    'bottom': {'style': 'SOLID', 'width': 1},
                    'left': {'style': 'SOLID', 'width': 1},
                    'right': {'style': 'SOLID', 'width': 1},
                    'innerHorizontal': {'style': 'SOLID', 'width': 1},
                    'innerVertical': {'style': 'SOLID', 'width': 1}
                }
            }
        ]

        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={'requests': requests}
        ).execute()

    def update_data(self, range_name, values):
        body = {'values': values}
        self.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f"{self.sheet_name}!{range_name}",
            body=body,
            valueInputOption='USER_ENTERED'
        ).execute()
