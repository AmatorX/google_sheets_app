
from collections import defaultdict
from datetime import date
from sheets.base_sheet import BaseTable
from buildings.models import WorkEntry, Material



class Sheet1Table(BaseTable):
    def __init__(self, obj):
        self.obj = obj
        self.sh_url = obj.sh_url
        super().__init__(obj=obj, sheet_name="Sheet1")

    def get_start_budget(self):
        """
        Возвращает бюджет на первую доступную дату текущего месяца.
        Если таких записей нет — возвращает исходный total_budget.
        """
        today = date.today()
        first_day = today.replace(day=1)

        # Первая запись за текущий месяц
        entry = self.obj.budget_history.filter(date__gte=first_day).order_by('date').first()

        if entry:
            print(f"Функция get_start_budget вернула {entry.current_budget}")
            return entry.current_budget
        return self.obj.total_budget

    def get_worker_month_salary(self, worker):
        """
        Возвращает зарплату одного работника за текущий месяц.
        """
        today = date.today()
        first_day = today.replace(day=1)

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            worker=worker,
            date__gte=first_day,
            date__lte=today
        )

        if worker.salary is None:
            return 0.0

        total_hours = sum(float(entry.worked_hours) for entry in entries)
        salary = round(total_hours * worker.salary, 2)
        return salary

    def get_all_workers_month_salary(self):
        """
        Возвращает общую зарплату за текущий месяц по объекту.
        """
        workers = self.obj.workers.all()
        total_salary = sum(self.get_worker_month_salary(worker) for worker in workers)
        return round(total_salary, 2)

    def get_worker_month_earned(self, worker):
        """
        Возвращает заработок работника за установленный материал за текущий месяц.
        """
        today = date.today()
        first_day = today.replace(day=1)

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            worker=worker,
            date__gte=first_day,
            date__lte=today
        )

        earned = 0.0
        for entry in entries:
            materials = entry.get_materials_used()
            for material_name, quantity in materials.items():
                try:
                    material = Material.objects.get(name=material_name)
                    price = material.price or 0
                    earned += float(quantity) * price
                except Material.DoesNotExist:
                    print(f"Материал '{material_name}' не найден. Пропускаем.")
                    continue
        return round(earned, 2)

    def get_month_earned(self):
        """
        Возвращает общую сумму заработка всех работников за текущий месяц по установленным материалам.
        """
        workers = self.obj.workers.all()
        total_earned = sum(self.get_worker_month_earned(worker) for worker in workers)
        return round(total_earned, 2)

    def get_summary_rows(self):
        """
        Возвращает сводные строки для записи в таблицу:
        1. Название месяца и года
        2. Bldg cost (начальный бюджет)
        3. Salary (зарплата за месяц)
        4. Earned (заработано за материалы)
        5. Result = Earned - Salary
        6. Progress = Earned / Bldg cost * 100
        """
        today = date.today()
        month_year = today.strftime("%B %Y")

        start_budget = self.get_start_budget()
        salary = self.get_all_workers_month_salary()
        earned = self.get_month_earned()

        result = earned - salary
        progress = (earned / start_budget * 100) if start_budget else 0

        return [
            [month_year],
            ['Bldg cost', start_budget],
            ['Salary', salary],
            ['Earned', earned],
            ['Result', result],
            ['Progress', round(progress, 2)]
        ]

    def write_summary(self):
        """
        Вставляет сводную таблицу: две пустые строки + блок данных.
        Затем применяет стили к вставленному блоку.
        """
        # Получаем текущую последнюю строку в колонке B
        start_row = self.get_last_non_empty_row(column_letter='A') + 1

        # Две пустые строки
        empty_rows = [[''], ['']]

        # Сводные данные
        summary_rows = self.get_summary_rows()

        # Общие данные для записи
        all_rows = empty_rows + summary_rows

        # Вычисляем диапазон, например 'A15'
        start_cell = f"A{start_row}"

        # Запись в таблицу
        self.update_data(start_cell, all_rows)

        # Применяем стили только к summary_rows (без пустых строк)
        style_start_row = start_row + len(empty_rows)
        style_end_row = style_start_row + len(summary_rows) - 1
        self.apply_styles(
            start_row=style_start_row,
            end_row=style_end_row,
            start_col=0,
            end_col=2
        )

        print(f"Сводная таблица записана, стили применены с {style_start_row} по {style_end_row - 1}")

    def get_workers_summary_rows(self):
        """
        Возвращает список строк:
        - две пустые строки,
        - строка с месяцем и годом,
        - строки: [имя работника, заработано / часы]
        — только для тех, кто оставлял WorkEntry в текущем месяце.
        """
        today = date.today()
        first_day = today.replace(day=1)

        # Все записи за текущий месяц по объекту
        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            date__gte=first_day,
            date__lte=today
        )

        worker_hours = defaultdict(float)
        for entry in entries:
            worker_hours[entry.worker] += float(entry.worked_hours)

        rows = []
        for worker, hours in worker_hours.items():
            if hours == 0:
                continue
            earned = self.get_worker_month_earned(worker)
            value = round(earned / hours, 2)
            rows.append([worker.name, value])

        # Добавляем две пустые строки и строку с месяцем
        month_year = [[today.strftime("%B %Y"), "Earned per hour"]]
        all_rows = [[''], ['']] + month_year + rows
        print(f"Строки для записи функции get_workers_summary_rows: {all_rows}")
        return all_rows

    def write_summary_workers(self):
        """
        Вставляет блок данных по работникам: две пустые строки, строка с месяцем,
        затем строки [имя работника, заработано / часы].
        Затем применяет стили к блоку с работниками (кроме пустых строк).
        """
        # Получаем текущую последнюю строку в колонке B
        start_row = self.get_last_non_empty_row(column_letter='A') + 1

        # Все строки, включая две пустых, месяц и список работников
        worker_rows = self.get_workers_summary_rows()

        # Вычисляем начальную ячейку, например "A27"
        start_cell = f"A{start_row}"

        # Запись данных
        self.update_data(start_cell, worker_rows)

        # Применяем стили ко всем строкам, кроме первых двух пустых
        style_start_row = start_row + 2
        style_end_row = start_row + len(worker_rows) - 1

        self.apply_styles(
            start_row=style_start_row,
            end_row=style_end_row,
            start_col=0,
            end_col=2
        )

        print(f"Сводка по работникам записана, стили применены с {style_start_row} по {style_end_row}")

    def get_total_materials_usage(self):
        """
        Возвращает словарь {название_материала: общее_количество}
        по всем WorkEntry для текущего объекта.
        """
        entries = WorkEntry.objects.filter(build_object=self.obj)
        total_usage = defaultdict(float)

        for entry in entries:
            try:
                materials = entry.get_materials_used()
            except Exception as e:
                print(f"Ошибка при чтении материалов из WorkEntry(id={entry.id}): {e}")
                continue

            for name, quantity in materials.items():
                try:
                    total_usage[name] += float(quantity)
                except (TypeError, ValueError) as e:
                    print(f"Ошибка при обработке количества материала '{name}' в WorkEntry(id={entry.id}): {e}")
        print(f"Объект {self.obj.name} использовано материалов {dict(total_usage)}")
        return dict(total_usage)


    def write_materials_summary(self, start_row=None):
        """
        Записывает в таблицу блок с суммой использованных материалов:
        Колонка 'Material' и колонка 'Quantity' рядом с ней.
        Если start_row не указан — записывает после основной KPI-таблицы.
        """
        summary = self.get_total_materials_usage()
        if not summary:
            print("Нет данных для записи материалов.")
            return

        # Формируем данные для записи
        materials = list(summary.items())
        data = [["Material", "Total usage"]] + [[name, round(qty, 2)] for name, qty in materials]

        # Если не указан старт, ставим под основной таблицей
        if start_row is None:
            start_row = self.get_last_non_empty_row() + 3

        cell = f"A{start_row}"
        self.update_data(cell, data)

        # Применяем стили
        self.apply_styles(
            start_row=start_row,
            end_row=start_row + len(data) - 1,
            start_col=0,
            end_col=2,
        )

        print(f"Сводка по материалам записана начиная с {cell}")

