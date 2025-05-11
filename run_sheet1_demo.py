import os
import django

# Настройка Django
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tsa.settings")
django.setup()

from buildings.models import BuildObject
from sheets.sheet1 import Sheet1Table
def main():
    build_objects = BuildObject.objects.all()

    if not build_objects.exists():
        print("Нет доступных объектов.")
        return

    for obj in build_objects:
        print(f"\n--- Объект: {obj.name} ---")
        try:
            table = Sheet1Table(obj)
            table.write_materials_summary()
        except Exception as e:
            print(f"Ошибка при обработке объекта '{obj.name}': {e}")

if __name__ == "__main__":
    main()
