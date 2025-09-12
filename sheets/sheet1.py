from calendar import monthrange
from collections import defaultdict
from datetime import date
from time import sleep

from sheets.base_sheet import BaseTable
from tsa_app.models import WorkEntry, Material


class Sheet1Table(BaseTable):
    def __init__(self, obj, work_spec=None):
        self.obj = obj
        self.sh_url = obj.sh_url
        self.work_spec = work_spec
        super().__init__(obj=obj, sheet_name="Sheet1")

    def get_worker_month_salary(self, worker, month=None, year=None):
        """
        Возвращает зарплату работника за указанный месяц.

        :param worker: Экземпляр Worker.
        :param month: Месяц (1–12). По умолчанию текущий.
        :param year: Год. По умолчанию текущий.
        :return: Общая сумма зарплаты.
        """
        today = date.today()
        if month is None: month = today.month
        if year is None: year = today.year
        first_day = date(year, month, 1)
        last_day = date(year, month, monthrange(year, month)[1])

        if worker.salary is None:
            return 0.0

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            worker=worker,
            date__range=(first_day, last_day)
        )
        total_hours = sum(float(entry.worked_hours) for entry in entries)
        return round(total_hours * worker.salary, 2)

    def get_all_workers_month_salary(self, month=None, year=None):
        """
        Возвращает суммарную зарплату всех работников объекта за указанный месяц.

        :param month: Месяц (1–12).
        :param year: Год.
        :return: Общая сумма зарплат.
        """
        workers = self.obj.workers.all()
        return round(sum(self.get_worker_month_salary(w, month, year) for w in workers), 2)

    def get_worker_month_earned(self, worker, month=None, year=None):
        """
        Возвращает сумму, заработанную работником за установку материалов за месяц.

        :param worker: Экземпляр Worker.
        :param month: Месяц.
        :param year: Год.
        :return: Сумма заработка за материалы.
        """
        today = date.today()
        if month is None: month = today.month
        if year is None: year = today.year
        first_day = date(year, month, 1)
        last_day = date(year, month, monthrange(year, month)[1])

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            worker=worker,
            date__range=(first_day, last_day)
        )

        earned = 0.0
        for entry in entries:
            materials = entry.get_materials_used()
            for name, qty in materials.items():
                try:
                    material = Material.objects.get(name=name)
                    earned += float(qty) * (material.price or 0)
                except Material.DoesNotExist:
                    continue
        return round(earned, 2)

    def get_month_earned(self, month=None, year=None):
        """
        Считает общую сумму заработка за установку материалов всеми работниками.

        :param month: Месяц.
        :param year: Год.
        :return: Общий заработок.
        """
        workers = self.obj.workers.all()
        return round(sum(self.get_worker_month_earned(w, month, year) for w in workers), 2)

    def get_summary_rows(self, month=None, year=None):
        """
        Формирует строки общей статистики: бюджет, зарплаты, заработано и прогресс.

        :param month: Месяц.
        :param year: Год.
        :return: Список строк для записи или None, если нет данных.
        """
        today = date.today()
        if month is None: month = today.month
        if year is None: year = today.year
        month_year = date(year, month, 1).strftime("%B %Y")

        budget = self.obj.total_budget
        salary = self.get_total_salary(month, year)
        if salary:
            earned = self.get_total_materials_income(month, year)
            result = earned - salary
            progress = f"{round((earned / budget * 100), 2) if budget else 0}%"
            return [
                [month_year],
                ['Bldg cost', budget],
                ['Salary', salary],
                ['Earned', earned],
                ['Result', round(result, 2)],
                ['Progress', progress]
            ]
        return None

    # def get_summary_rows(self, month=None, year=None):
    #     work_spec = self.work_spec
    #     print("Запуск get_summary_rows")
    #     print(f"Передан {work_spec=}, type: {type(work_spec)}")
    #     print(f"get_summary_rows work_spec id: {id(work_spec)}")
    #     today = date.today()
    #     if month is None: month = today.month
    #     if year is None: year = today.year
    #
    #     month_year = date(year, month, 1).strftime("%B %Y")
    #
    #     entries = WorkEntry.objects.filter(
    #         build_object=self.obj,
    #         date__year=year,
    #         date__month=month
    #     )
    #     print(f"get_summary_rows получены entries: {entries}")
    #
    #     if work_spec:
    #         entries = entries.filter(worker__work_specialization=work_spec)
    #         print(f"После фильтрации по спецификации entries: {entries}")
    #
    #     if not entries.exists():
    #         return None
    #
    #     total_hours = 0
    #     total_earned = 0
    #     print(f"<<<<<< entries {entries} >>>>>>")
    #     for entry in entries:
    #         print(f"Работа: {entry.worker.name}, часы: {entry.worked_hours}, материалы_raw: {entry.materials_used}")
    #         hours = entry.get_worked_hours_as_float()
    #         salary = entry.worker.salary or 0
    #
    #         total_hours += hours
    #         total_earned_entry = 0
    #
    #         materials = entry.get_materials_used()
    #         for _, material_info in materials.items():
    #             total_earned_entry += material_info["quantity"] * material_info["unit_price"]
    #
    #         total_earned += total_earned_entry
    #
    #     total_salary = sum(entry.get_worked_hours_as_float() * (entry.worker.salary or 0) for entry in entries)
    #     result = total_earned - total_salary
    #     budget = self.obj.total_budget
    #     progress = f"{round((total_earned / budget * 100), 2) if budget else 0}%"
    #     return [
    #         [month_year],
    #         ['Bldg cost', budget],
    #         ['Salary', round(total_salary, 2)],
    #         ['Earned', round(total_earned, 2)],
    #         ['Result', round(result, 2)],
    #         ['Progress', progress]
    #     ]

    def write_summary(self, month=None, year=None):
        """
        Записывает строки общей статистики в таблицу (две пустые строки + блок).

        :param month: Месяц.
        :param year: Год.
        """
        start_row = self.get_last_non_empty_row(column_letter='B') + 1
        empty_rows = [[' '], [' ']]
        summary_rows = self.get_summary_rows(month, year)
        print(f"Функция write_summary для записи переданы данные: {summary_rows}")
        if summary_rows:
            all_rows = empty_rows + summary_rows
            print(f"write_summary {self.obj.name}: {all_rows} -----")
            self.update_data(f"A{start_row}", all_rows)
            self.apply_styles(
                start_row=start_row + 2,
                end_row=start_row + len(all_rows) - 1,
                start_col=0,
                end_col=2
            )

    def get_workers_summary_rows(self, month=None, year=None):
        """
        Возвращает строки по работникам: имя и заработок (за установка материалов) за час за месяц.

        :param month: Месяц.
        :param year: Год.
        :return: Список строк или None, если нет записей.
        """
        today = date.today()
        if month is None: month = today.month
        if year is None: year = today.year
        first_day = date(year, month, 1)
        last_day = date(year, month, monthrange(year, month)[1])

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            date__range=(first_day, last_day)
        )
        if not entries.exists():
            return None

        worker_hours = defaultdict(float)
        for entry in entries:
            worker_hours[entry.worker] += float(entry.worked_hours)

        rows = []
        for worker, hours in worker_hours.items():
            if hours == 0:
                continue
            earned = self.get_worker_month_earned(worker, month, year)
            rows.append([worker.name, round(earned / hours, 2)])

        month_year = [[date(year, month, 1).strftime("%B %Y"), "Earned per hour"]]
        return [[' '], [' ']] + month_year + rows

    def write_summary_workers(self, month=None, year=None):
        """
        Пишет блок по работникам (заработано/часы) в таблицу.
        Предварительно проверяет и удаляет если за записываемый месяц уже есть таблица

        :param month: Месяц.
        :param year: Год.
        :return: Список строк без заголовков, либо None.
        """
        self.remove_existing_month_data_if_exists(month, year)
        start_row = self.get_last_non_empty_row(column_letter='B') + 1
        rows = self.get_workers_summary_rows(month, year)
        print(f"Функция write_summary_workers для записи в таблицу пользователей переданы данные: {rows}")
        if rows:
            self.update_data(f"A{start_row}", rows)
            sleep(2)
            self.apply_styles(
                start_row=start_row + 2,
                end_row=start_row + len(rows) - 1,
                start_col=0,
                end_col=2
            )
            sleep(2)
            print(f"write_summary_workers {self.obj.name}: {rows}")
            return rows[3:]
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

    # def get_total_salary(self, month=None, year=None):
    #     """
    #     Возвращает общую зарплату по WorkEntry текущего объекта за месяц.
    #
    #     :param month: Месяц.
    #     :param year: Год.
    #     :return: Общая сумма зарплат.
    #     """
    #     today = date.today()
    #     if month is None: month = today.month
    #     if year is None: year = today.year
    #     first_day = date(year, month, 1)
    #     last_day = date(year, month, monthrange(year, month)[1])
    #     entries = WorkEntry.objects.filter(
    #         build_object=self.obj,
    #         date__range=(first_day, last_day)
    #     ).select_related('worker')
    #
    #     total = 0.0
    #     for entry in entries:
    #         rate = entry.worker.salary or 0
    #         hours = entry.get_worked_hours_as_float()
    #         total += rate * hours
    #     return round(total, 2)

    def get_total_salary(self, month=None, year=None):
        """
        Возвращает общую зарплату по WorkEntry текущего объекта за месяц с учётом work_spec.

        :param month: Месяц.
        :param year: Год.
        :return: Общая сумма зарплат.
        """
        today = date.today()
        if month is None: month = today.month
        if year is None: year = today.year

        first_day = date(year, month, 1)
        last_day = date(year, month, monthrange(year, month)[1])

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            date__range=(first_day, last_day)
        ).select_related('worker')

        if self.work_spec:
            entries = entries.filter(worker__work_specialization=self.work_spec)

        total = 0.0
        for entry in entries:
            if not entry.worker:
                continue  # пропускаем, если worker удалён
            rate = entry.worker.salary or 0
            hours = entry.get_worked_hours_as_float()
            total += rate * hours

        return round(total, 2)

    # def get_total_materials_income(self, month=None, year=None):
    #     """
    #     Возвращает доход от установки материалов по WorkEntry объекта за месяц.
    #
    #     :param month: Месяц.
    #     :param year: Год.
    #     :return: Общий доход.
    #     """
    #     today = date.today()
    #     if month is None: month = today.month
    #     if year is None: year = today.year
    #     first_day = date(year, month, 1)
    #     last_day = date(year, month, monthrange(year, month)[1])
    #
    #     entries = WorkEntry.objects.filter(
    #         build_object=self.obj,
    #         date__range=(first_day, last_day)
    #     )
    #
    #     total = 0.0
    #     cache = {}
    #     for entry in entries:
    #         for name, qty in entry.get_materials_used().items():
    #             try:
    #                 qty = float(qty)
    #             except Exception:
    #                 qty = 0
    #             if name not in cache:
    #                 material = Material.objects.filter(name__iexact=name).first()
    #                 cache[name] = material.price if material and material.price else 0
    #             total += cache[name] * qty
    #     return round(total, 2)

    def get_total_materials_income(self, month=None, year=None):
        """
        Возвращает доход от установки материалов по WorkEntry объекта за месяц с учётом work_spec.

        :param month: Месяц.
        :param year: Год.
        :return: Общий доход.
        """
        today = date.today()
        if month is None: month = today.month
        if year is None: year = today.year

        first_day = date(year, month, 1)
        last_day = date(year, month, monthrange(year, month)[1])

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            date__range=(first_day, last_day)
        ).select_related('worker')

        if self.work_spec:
            entries = entries.filter(worker__work_specialization=self.work_spec)

        total = 0.0
        cache = {}
        for entry in entries:
            materials = entry.get_materials_used()
            for name, qty in materials.items():
                try:
                    qty = float(qty)
                except Exception:
                    qty = 0
                if name not in cache:
                    material = Material.objects.filter(name__iexact=name).first()
                    cache[name] = material.price if material and material.price else 0
                total += cache[name] * qty

        return round(total, 2)

    def remove_existing_month_data_if_exists(self, month=None, year=None):
        """
        Удаляет из таблицы блок строк, соответствующий переданному месяцу и году.

        :param month: Месяц.
        :param year: Год.
        """
        today = date.today()
        if month is None: month = today.month
        if year is None: year = today.year
        month_year = date(year, month, 1).strftime("%B %Y")
        print(f"remove_existing_month_data_if_exists ищем таблицы для {month_year}")
        range_name = f"{self.sheet_name}!A1:A"
        response = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=range_name
        ).execute()
        values = response.get('values', [])
        start_row = None
        for idx, row in enumerate(values):
            if row and row[0].strip() == month_year:
                start_row = idx
                print(f"remove_existing_month_data_if_exists совпадение для таблиц месяца {month_year} найдено в {idx}")
                break
        if start_row is None:
            return
        end_row = start_row
        for i in range(start_row + 1, len(values)):
            if values[i] and values[i][0]:
                end_row = i
            else:
                break
        delete_start = start_row
        delete_end = end_row + 2
        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={
                "requests": [{
                    "deleteDimension": {
                        "range": {
                            "sheetId": self.sheet_id,
                            "dimension": "ROWS",
                            "startIndex": delete_start,
                            "endIndex": delete_end + 1
                        }
                    }
                }]
            }
        ).execute()
        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={
                "requests": [{
                    "insertDimension": {
                        "range": {
                            "sheetId": self.sheet_id,
                            "dimension": "ROWS",
                            "startIndex": delete_start,
                            "endIndex": delete_end + 1
                        }
                    }
                }]
            }
        ).execute()
