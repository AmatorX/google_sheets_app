from calendar import monthrange
from collections import defaultdict
from datetime import date
from sheets.base_sheet import BaseTable
from tsa_app.models import WorkEntry, Material



class Sheet1Table(BaseTable):
    def __init__(self, obj):
        self.obj = obj
        self.sh_url = obj.sh_url
        super().__init__(obj=obj, sheet_name="Sheet1")

    def get_worker_month_salary(self, worker):
        """
        Возвращает зарплату одного работника за текущий месяц.
        """
        #######################################################################################################
        # year = 2025
        # month = 5
        #
        # first_day = date(year, month, 1)
        # today = date(year, month, monthrange(year, month)[1])
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
        print(f"get_all_workers_month_salary работники {workers}")
        total_salary = sum(self.get_worker_month_salary(worker) for worker in workers)
        print(f"get_all_workers_month_salary объект {self.obj.name} общая зп {round(total_salary, 2)}")
        return round(total_salary, 2)

    def get_worker_month_earned(self, worker, month=None, year=None):
        """
        Возвращает заработок работника за установленный материал за указанный месяц.
        Если month и year не заданы, используется текущий месяц.
        """
        ####################################################################################################
        print(f"Функция get_worker_month_earned для работника {worker} запущена")
        today = date.today()
        if month is None:
            month = today.month
        if year is None:
            year = today.year

        first_day = date(year, month, 1)
        last_day = date(year, month, monthrange(year, month)[1])

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            worker=worker,
            date__gte=first_day,
            date__lte=last_day
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
        print(f"get_month_earned объект {self.obj.name} заработок работников {round(total_earned, 2)}")
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
        # ######################################################################### хардкод для мая
        # #
        # year = 2025
        # month = 5
        # today = date(year, month, monthrange(year, month)[1])
        # month_year = today.strftime("%B %Y")

        # start_budget = self.get_start_budget()
        budget = self.obj.total_budget
        salary = self.get_total_salary()
        if salary:
            earned = self.get_total_materials_income()

            result = earned - salary
            progress = round(((earned / budget * 100) if budget else 0), 2)
            progress = str(progress)+ '%'

            return [
                [month_year],
                ['Bldg cost', budget],
                ['Salary', salary],
                ['Earned', earned],
                ['Result', result],
                ['Progress', progress]
            ]
        return None

    def write_summary(self):
        """
        Вставляет сводную таблицу: две пустые строки + блок данных.
        Затем применяет стили к вставленному блоку.
        """
        # Получаем текущую последнюю строку в колонке B
        start_row = self.get_last_non_empty_row(column_letter='B') + 1

        # Две пустые строки
        empty_rows = [[' '], [' ']]

        # Сводные данные
        summary_rows = self.get_summary_rows()
        if summary_rows:

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

        # ############################################################################# для апреля и марта хардкод
        # year = 2025
        # month = 5  # или 4 для апреля
        # first_day = date(year, month, 1)
        # today = date(year, month, monthrange(year, month)[1])

        # Все записи за текущий месяц по объекту
        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            date__gte=first_day,
            date__lte=today
        )
        if not entries.exists():  # Если записей нет — сразу выходим
            return None

        worker_hours = defaultdict(float)
        for entry in entries:
            worker_hours[entry.worker] += float(entry.worked_hours)
            print(f"{entry.worker.name} worker_hours -> {worker_hours}")

        rows = []
        for worker, hours in worker_hours.items():
            if hours == 0:
                continue
            earned = self.get_worker_month_earned(worker)
            value = round(earned / hours, 2)
            print(f"Работник {worker.name} заработано {earned} часов {hours} реальная зп {value}")
            rows.append([worker.name, value])

        # Добавляем две пустые строки и строку с месяцем
        month_year = [[today.strftime("%B %Y"), "Earned per hour"]]
        all_rows = [[' '], [' ']] + month_year + rows
        print(f"Строки для записи функции get_workers_summary_rows: {all_rows}")
        return all_rows

    def write_summary_workers(self):
        """
        Вставляет блок данных по работникам: две пустые строки, строка с месяцем,
        затем строки [имя работника, заработано / часы].
        Затем применяет стили к блоку с работниками (кроме пустых строк).
        """
        # Получаем текущую последнюю строку в колонке B
        self.remove_existing_month_data_if_exists()
        start_row = self.get_last_non_empty_row(column_letter='B') + 1

        # Все строки, включая две пустых, месяц и список работников
        worker_rows = self.get_workers_summary_rows()

        if worker_rows:
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
            return worker_rows[3:]
        return None

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

    # def get_total_salary(self):
    #     """
    #     Возвращает общую сумму зарплаты для всех работников, оставивших отчеты по объекту
    #     за текущий месяц (с 1-го числа по сегодняшний день).
    #     Расчёт: почасовая оплата * количество отработанных часов.
    #     """
    #
    #     today = date.today()
    #     first_day_of_month = today.replace(day=1)
    #
    #     entries = WorkEntry.objects.filter(
    #         build_object=self.obj,
    #         date__range=(first_day_of_month, today)
    #     ).select_related('worker')
    #
    #     total = 0.0
    #     for entry in entries:
    #         salary_per_hour = entry.worker.salary or 0
    #         worked_hours = entry.get_worked_hours_as_float()
    #         total += salary_per_hour * worked_hours
    #
    #     return round(total, 2)


    def get_total_salary(self):
        """
        Вычисляет суммарную заработную плату всех работников,
        оставивших отчёты по данному объекту за текущий месяц.

        Расчёт производится по формуле:
            зарплата = сумма (почасовая ставка работника * количество отработанных часов)

        Только те WorkEntry, у которых дата входит в текущий месяц (с 1-го числа до сегодня), учитываются.

        :return: Общая сумма заработной платы за текущий месяц (округлена до 2 знаков после запятой).
        :rtype: float
        """
        today = date.today()
        first_day_of_month = today.replace(day=1)

        ######################################################################## хардкод для мая

        # year = 2025
        # month = 5
        #
        # first_day_of_month = date(year, month, 1)
        # today = date(year, month, monthrange(year, month)[1])

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            date__range=(first_day_of_month, today)
        ).select_related('worker')

        total = 0.0
        for entry in entries:
            salary_per_hour = entry.worker.salary or 0
            worked_hours = entry.get_worked_hours_as_float()
            total += salary_per_hour * worked_hours

        return round(total, 2)

    def get_total_materials_income(self):
        """
        Вычисляет общий доход от установки материалов по объекту за текущий месяц.

        Проходится по всем WorkEntry текущего месяца, связанным с объектом,
        извлекает материалы из поля `materials_used` (в виде словаря: {название: количество}),
        находит цену каждого материала (если указана) и считает:
            доход = сумма (цена материала * количество)

        Материалы ищутся среди объектов Material, связанных с текущим BuildObject.
        Если материал с таким названием не найден или не указана цена — берётся 0.

        :return: Общая сумма дохода за установку материалов за текущий месяц (округлена до 2 знаков после запятой).
        :rtype: float
        """
        today = date.today()
        first_day_of_month = today.replace(day=1)
        # ############################################################################# хардкод для мая
        # year = 2025
        # month = 5
        #
        # first_day_of_month = date(year, month, 1)
        # today = date(year, month, monthrange(year, month)[1])

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            date__range=(first_day_of_month, today)
        )

        total = 0.0
        material_cache = {}  # Кэш цен материалов по названию
        print(f"Объект: {self.obj.name}")
        print(f"Всего записей: {entries.count()}")
        for entry in entries:
            print(f"Дата: {entry.date}")
            print(f"Материалы: {entry.get_materials_used()}")
            materials_dict = entry.get_materials_used()
            for mat_name, quantity in materials_dict.items():
                try:
                    quantity = float(quantity)
                except (ValueError, TypeError):
                    quantity = 0  # если не удается привести — считаем как 0

                if mat_name not in material_cache:
                    material_obj = Material.objects.filter(
                        name__iexact=mat_name
                    ).first()
                    material_cache[mat_name] = material_obj.price if material_obj and material_obj.price else 0

                total += material_cache[mat_name] * quantity
        return round(total, 2)


    def remove_existing_month_data_if_exists(self):
        sheet_name = self.sheet_name
        range_name = f"{sheet_name}!A1:A"
        month_year = date.today().strftime("%B %Y")

        # Получаем значения из колонки A
        response = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=range_name
        ).execute()

        col_values = response.get('values', [])

        # Поиск строки со значением текущий месяц год
        start_row = None
        for idx, row in enumerate(col_values):
            if row and row[0].strip() == month_year:
                start_row = idx
                break

        if start_row is None:
            print("Текущий месяц не найден — ничего не удаляем.")
            return
        # Определяем границу таблицы — пока есть непустые строки
        end_row = start_row
        for i in range(start_row + 1, len(col_values)):
            if col_values[i] and col_values[i][0]:
                end_row = i
            else:
                break

        # Удалим строки от (start_row) до (end_row + 2)
        delete_start = max(0, start_row )
        delete_end = end_row + 2

        print(f"Удаляем строки с {delete_start + 1} по {delete_end + 1}")

        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={
                "requests": [
                    {
                        "deleteDimension": {
                            "range": {
                                "sheetId": self.sheet_id,
                                "dimension": "ROWS",
                                "startIndex": delete_start,
                                "endIndex": delete_end + 1  # endIndex не включительно
                            }
                        }
                    }
                ]
            }
        ).execute()
        print("Функция удаления существующего отчета за месяц. Удаление завершено.")
