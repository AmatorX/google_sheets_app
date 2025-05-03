import os
import django

# Устанавливаем переменную окружения для настроек Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'tsa.settings')
django.setup()

from sheets.results_sheet import ResultsTable
from buildings.models import BuildObject

if __name__ == "__main__":
    build_objects = BuildObject.objects.all()

    for build_object in build_objects:
        try:
            print(f"Обрабатываем объект: {build_object.name}")
            results_table = ResultsTable(obj=build_object)

            # Убедимся, что лист существует
            results_table.ensure_sheet_exists()

            # Получаем и строим таблицу для работников
            all_worker_data = results_table.build_worker_table()
            if all_worker_data:
                # Запишем данные в Google Sheets
                results_table.write_to_google_sheets(all_worker_data)

        except Exception as e:
            print(f"Ошибка при обработке объекта {build_object.name}: {e}\n")
