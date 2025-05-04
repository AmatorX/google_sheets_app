import os
import django

# Настройка Django
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tsa.settings")
django.setup()

from buildings.models import BuildObject
from sheets.sheet1 import DefaultSheetTable
def main():
    build_objects = BuildObject.objects.all()

    if not build_objects.exists():
        print("Нет доступных объектов.")
        return

    for obj in build_objects:
        print(f"\n--- Объект: {obj.name} ---")
        try:
            table = DefaultSheetTable(obj)
            # budget = table.get_start_budget()
            # print(f"Стартовый бюджет: {budget}")
            # salary = table.get_all_workers_month_salary()
            # earned = table.get_month_earned()
            table.write_summary()
            # workers_rows = table.get_workers_summary_rows()
            table.write_summary_workers()
        except Exception as e:
            print(f"Ошибка при обработке объекта '{obj.name}': {e}")

if __name__ == "__main__":
    main()
