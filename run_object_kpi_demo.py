import os
import django

# Настройка Django
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tsa.settings")
django.setup()

from tsa_app.models import BuildObject
from sheets.object_kpi_sheet import ObjectKPITable

def main():
    # build_objects = BuildObject.objects.all()
    build_objects = BuildObject.objects.filter(is_archived=False)

    if not build_objects.exists():
        print("Нет доступных объектов.")
        return

    for obj in build_objects:
        print(f"\n--- Объект: {obj.name} ---")
        try:
            table = ObjectKPITable(obj)
            # Записываем расчёты
            table.write_today_data()
            print("Данные успешно записаны.")
        except Exception as e:
            print(f"Ошибка при обработке объекта '{obj.name}': {e}")

if __name__ == "__main__":
    main()
