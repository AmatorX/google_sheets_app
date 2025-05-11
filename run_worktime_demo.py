
import os
import django

# Устанавливаем переменную окружения для настроек Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'tsa.settings')
django.setup()

from sheets.work_time_sheet import WorkTimeTable
# from buildings.models import BuildObjectfrom
from tsa_app.models import BuildObject
#
#
# if __name__ == "__main__":
#     # Выбираем нужный BuildObject (по id)
#     # build_object = BuildObject.objects.get(id=2)
#     build_objects = BuildObject.objects.all()
#
#     chunk_data = create_chunk_data(full_list=False)
#     print(f'Chunk data: {chunk_data}')
#     for build_object in build_objects:
#         try:
#             print(f'Работаем с объектом {build_object.name}')
#             worktime_table = WorkTimeTable(obj=build_object)
#
#             # Убедимся, что лист существует
#             worktime_table.ensure_sheet_exists()
#
#             # Удаляем старую таблицу, если есть
#             worktime_table.remove_existing_chunk_if_exists(chunk_data)
#
#             # Генерация строк для таблицы
#             chunk, rows_generated = worktime_table.generate_current_chunk_data_and_rows()
#
#             if rows_generated:
#                 worktime_table.create_table(chunk_data, rows_generated)
#             else:
#                 print(f'Данных для записи в Work Time для объекта {build_object.name} нет.\n')
#         except Exception as e:
#             print(f'Ошибка при обработке объекта {build_object.name}: {e}\n')


if __name__ == "__main__":
    build_objects = BuildObject.objects.all()
    for build_object in build_objects:
        try:
            print(f'Обрабатываем объект: {build_object.name}')
            worktime_table = WorkTimeTable(obj=build_object)
            worktime_table.update_current_chunk()
        except Exception as e:
            print(f'Ошибка: {e}\n')