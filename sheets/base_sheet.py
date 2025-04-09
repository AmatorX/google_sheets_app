import logging
from utils.service import get_service

logger = logging.getLogger(__name__)

class BaseTable:
    def __init__(self, spreadsheet_url, sheet_name):
        self.spreadsheet_url = spreadsheet_url
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
        return self.spreadsheet_url.split("/")[5]

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

    def get_last_non_empty_row(self):
        """Возвращает индекс последней непустой строки в столбце A"""
        try:
            range_name = f"{self.sheet_name}!B1:B"
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()
            values = result.get('values', [])
            return len(values) if values else 0
        except Exception as e:
            logger.error(f"Ошибка при получении последней непустой строки: {e}")
            return 0

    def apply_styles(self, start_row, end_row, start_col, end_col, color):
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
                    'cell': {'userEnteredFormat': {'backgroundColor': color}},
                    'fields': 'userEnteredFormat.backgroundColor'
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

