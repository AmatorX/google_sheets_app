import os
import django

# Устанавливаем переменную окружения для настроек Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'tsa.settings')
django.setup()

from sheets.photos_sheet import PhotosTable
from buildings.models import BuildObject

if __name__ == "__main__":
    build_objects = BuildObject.objects.all()

    for build_object in build_objects:
        try:
            print(f"Обрабатываем объект: {build_object.name}")
            photos_table = PhotosTable(build_object=build_object)

            # Убедимся, что лист существует
            photos_table.ensure_sheet_exists()

            # Создаём таблицы для воркеров, если их ещё нет
            photos_table.create_tables_for_missing_workers()

        except Exception as e:
            print(f"Ошибка при обработке объекта {build_object.name}: {e}\n")
