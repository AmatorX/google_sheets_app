from datetime import date, datetime
from decimal import Decimal


from sheets.base_sheet import BaseTable
from tsa_app.models import WorkEntry, Material, BuildBudgetHistory


class ObjectKPITable(BaseTable):
    """
    Класс для работы с листом 'Object KPI' в Google Таблице, связанной с объектом строительства.
    """

    def __init__(self, build_object):
        self.obj = build_object
        self.sheet_name = "Object KPI"
        self.current_budget = build_object.current_budget
        current_month = datetime.now().strftime("%B")
        self.header = [
            current_month,
            "Done today $",
            "Done today %",
            "Total left %",
            "Budget balance"
        ]
        self.header2 = [" ", " ", " ", " ", self.current_budget]
        self.sheet_range = f"{self.sheet_name}!A1"
        super().__init__(build_object, sheet_name=self.sheet_name)

    def get_done_today_dollars(self):
        """
        Возвращает общую стоимость установленных материалов (в долларах) за сегодня на объекте.
        """
        today = date.today()
        total = Decimal('0.0')

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            date=today
        )
        print(f"метод get_done_today_dollars entries: {entries}")

        for entry in entries:
            materials = entry.get_materials_used()
            for material_name, qty in materials.items():
                price = self.get_material_price(material_name)
                total += Decimal(qty) * Decimal(price)

        return float(total)

    def get_material_price(self, name):
        """
        Получает цену за единицу материала по названию.
        """
        try:
            material = Material.objects.get(name=name)
            return material.price
        except Material.DoesNotExist:
            print(f"Цена для материала '{name}' не найдена, возвращаем 0")
            return 0

    # def update_budget(self, done_today_dollars):
    #     """
    #     Обновляет бюджет, исходя из текущей стоимости работ.
    #     """
    #     new_budget = self.obj.current_budget - done_today_dollars
    #     self.obj.current_budget = new_budget
    #     self.obj.save()


    def ensure_headers(self):
        """
        Проверяет, есть ли в столбце A название текущего месяца.
        Если нет - добавляет заголовки.
        """
        current_month = datetime.now().strftime("%B")
        range_name = f"{self.sheet_name}!A1:A"
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=range_name
        ).execute()
        values = result.get('values', [])

        # Проверяем, есть ли название месяца в первой колонке
        month_in_column = any(
            row and row[0].strip() == current_month for row in values
        )

        if not month_in_column:
            self.write_headers([self.header, self.header2])

    def update_budget_history(self, new_budget):
        today = date.today()

        BuildBudgetHistory.objects.update_or_create(
            build_object=self.obj,
            date=today,
            defaults={'current_budget': new_budget}
        )

    def write_today_row(self):
        """
        Добавляет строку с расчётами за сегодня в таблицу и обновляет бюджет.
        """
        today = date.today()
        day_of_month = today.day

        done_today = self.get_done_today_dollars()
        budget_before = self.obj.current_budget

        new_budget = round(budget_before - done_today, 2)
        done_percent = (done_today / budget_before * 100) if budget_before else 0
        left_percent = (new_budget / budget_before * 100) if budget_before else 0

        row = [
            day_of_month,
            round(done_today, 2),
            round(done_percent, 2),
            round(left_percent, 2),
            round(new_budget, 2)
        ]
        self.write_rows([row])
        # Обновляем бюджет и сохраняем
        self.obj.current_budget = new_budget
        self.obj.save()
        self.update_budget_history(new_budget)

    def write_headers(self, headers: list[list]):
        """
        Записывает заголовки начиная с первой строки. Добавляет отступы (3 строки).
        """
        start_row = self.get_last_non_empty_row(column_letter='E') + 1
        range_name = f"A{start_row}"
        rows = [[" "]] * 3 + headers  # Отступ из 3 строк
        self.update_data(range_name, rows)

        # Применим стили к заголовкам
        start_row += 3 # Потому что добавили три пустые строки перед заголовками
        end_row = start_row + len(headers) -1
        start_col = 0
        end_col = len(headers[0]) if headers else 1

        self.apply_styles(start_row, end_row, start_col, end_col)
        print(f"Данные записаны и стилизованы: строки {start_row}-{end_row - 1}")

    def write_rows(self, rows: list[list]):
        """
        Добавляет строки данных ниже последней непустой строки в столбце B.
        """
        start_row = self.get_last_non_empty_row(column_letter='E') + 1

        range_name = f"A{start_row}"
        self.update_data(range_name, rows)

        # Применим стили
        end_row = start_row + len(rows) -1
        start_col = 0
        end_col = len(rows[0]) if rows else 1

        self.apply_styles(start_row, end_row, start_col, end_col)

        print(f"Данные записаны и стилизованы: строки {start_row}-{end_row - 1}")

    def write_today_data(self):
        """
        Создаёт лист (если отсутствует), заголовки (если не заданы) и записывает строку с сегодняшними KPI.
        """
        try:
            self.ensure_sheet_exists()
            self.ensure_headers()
            self.write_today_row()
            print(f"KPI успешно записаны для объекта '{self.obj.name}'")
        except Exception as e:
            print(f"Ошибка при записи KPI для объекта '{self.obj.name}': {e}")
