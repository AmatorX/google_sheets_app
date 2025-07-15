
from typing import List
from sheets.base_sheet import BaseTable
from tsa_app.models import Tool, ToolsSheet


class ToolsTable(BaseTable):
    def __init__(self, tools_sheet: ToolsSheet):
        self.tools_sheet = tools_sheet
        self.sh_url = tools_sheet.sh_url
        self.sheet_name = tools_sheet.name
        self.header = ["Tool name", "Tool ID", "Price", "Date of issue", "Assigned to"]
        super().__init__(obj=tools_sheet, sheet_name = self.sheet_name)


    def get_tools_data(self) -> List[List[str]]:
        """
        Получает все инструменты и форматирует строки для таблицы.
        """
        tools = Tool.objects.select_related("assigned_to").all()

        rows = []
        for tool in tools:
            rows.append([
                tool.name,
                tool.tool_id,
                str(tool.price),
                tool.date_of_issue.strftime("%Y-%m-%d") if tool.date_of_issue else '',
                str(tool.assigned_to) if tool.assigned_to else '',
            ])
        return rows

    def build_table(self) -> List[List[str]]:
        """
        Собирает таблицу с заголовком и данными.
        """
        return [self.header] + self.get_tools_data()

    def update_tools_table(self):
        """
        Создаёт или обновляет таблицу инструментов на листе и применяет стили.
        """
        self.ensure_sheet_exists()
        self.clear_sheet()
        values = self.build_table()

        # Подготовка тела запроса на запись данных
        body = {
            "values": values
        }

        # Записываем данные начиная с ячейки A1
        self.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f"{self.sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()

        # Применяем стили ко всей таблице
        total_rows = len(values)
        total_cols = len(values[0]) if values else 0
        if total_rows > 0 and total_cols > 0:
            self.apply_styles(
                start_row=1,
                end_row=total_rows,
                start_col=0,
                end_col=total_cols,
                bold=True,
            )
