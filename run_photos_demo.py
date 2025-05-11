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
            photos_table = PhotosTable(obj=build_object)

            # Создаём таблицы для воркеров, если их ещё нет, записываем в гугл шитс
            photos_table.write_missing_worker_tables()

        except Exception as e:
            print(f"Ошибка при обработке объекта {build_object.name}: {e}\n")
