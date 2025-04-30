import os
import django

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tsa.settings")
django.setup()

from sheets.tools_sheet import ToolsTable
from buildings.models import ToolsSheet, Tool  # Добавлен Tool

if __name__ == "__main__":
    try:
        tools_sheet = ToolsSheet.get_solo()  # Используем singleton
        if tools_sheet is None:
            print("Нет объекта ToolsSheet в базе данных.")
        else:
            print(f"Обрабатываем лист: {tools_sheet.name}")
            print(f"URL листа {tools_sheet.sh_url}")

            table = ToolsTable(tools_sheet=tools_sheet)


            all_tools = Tool.objects.all()

            if all_tools.exists():
                table.create_or_update_sheet()  # Обновление таблицы
            else:
                print("Нет инструментов, связанных с этим листом.")
    except Exception as e:
        # Если tools_sheet был найден, то выводим его имя, иначе выводим N/A
        print(f"Ошибка при обработке листа {tools_sheet.name if tools_sheet else 'N/A'}: {e}")
