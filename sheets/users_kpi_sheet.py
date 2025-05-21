from collections import defaultdict
from datetime import date, timedelta

from tsa_app.models import WorkEntry, Material
from sheets.base_sheet import BaseTable



class UsersKPITable(BaseTable):
    """
    Класс для создания и управления таблицей KPI работников.
    Таблица перезаписывается каждый день с 1-го числа текущего месяца.
    """

    def __init__(self, obj):
        self.obj = obj
        self.sheet_name = "Users KPI"
        self.sh_url = self.obj.sh_url
        super().__init__(obj, sheet_name=self.sheet_name)

    def get_kpi_rows(self):
        """
        Возвращает список строк:
        [дата, имя1, имя2, ..., per_day] — заголовок
        [дата1, earned1, earned2, ..., per_day_sum] — данные по дням
        ...
        ['Total', sum1, sum2, ..., total_sum] — итоги по каждому пользователю и всего
        """
        today = date.today()
        first_day = today.replace(day=1)

        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            date__gte=first_day,
            date__lte=today
        )

        workers = list({entry.worker for entry in entries})
        workers.sort(key=lambda w: w.name)

        day_data = defaultdict(dict)
        for entry in entries:
            earned = self.get_worker_day_earned(entry.worker, entry.date)
            day_data[entry.date][entry.worker] = earned

        # Добавляем две пустые строки перед заголовком
        rows = [[""], [""]]
        header = ["Date"] + [worker.name for worker in workers] + ["Per day"]
        rows.append(header)

        totals = [0 for _ in workers]
        grand_total = 0

        current = first_day
        while current <= today:
            row = [current.strftime("%Y-%m-%d")]
            day_total = 0
            for i, worker in enumerate(workers):
                earned = day_data.get(current, {}).get(worker, "")
                if earned != "":
                    totals[i] += earned
                    day_total += earned
                    row.append(round(earned, 2))
                else:
                    row.append("")
            row.append(round(day_total, 2) if day_total else "")
            grand_total += day_total
            rows.append(row)
            current += timedelta(days=1)

        total_row = ["Total"] + [round(t, 2) for t in totals] + [round(grand_total, 2)]
        rows.append(total_row)

        return rows

    # def get_kpi_rows(self):
    #     """
    #     Возвращает список строк:
    #     [дата, имя1, имя2, ...] — заголовок
    #     [дата1, earned1, earned2, ...] — данные по дням
    #     ...
    #     ['Total', sum1, sum2, ...] — итоги по каждому пользователю
    #     """
    #     today = date.today()
    #     first_day = today.replace(day=1)
    #
    #     entries = WorkEntry.objects.filter(
    #         build_object=self.obj,
    #         date__gte=first_day,
    #         date__lte=today
    #     )
    #
    #     workers = list({entry.worker for entry in entries})
    #     workers.sort(key=lambda w: w.name)
    #
    #     day_data = defaultdict(dict)
    #     for entry in entries:
    #         earned = self.get_worker_day_earned(entry.worker, entry.date)
    #         day_data[entry.date][entry.worker] = earned
    #
    #     # Добавляем две пустые строки перед заголовком
    #     rows = [[""], [""]]
    #     header = ["Date"] + [worker.name for worker in workers]
    #     rows.append(header)
    #
    #     totals = [0 for _ in workers]
    #
    #     current = first_day
    #     while current <= today:
    #         row = [current.strftime("%Y-%m-%d")]
    #         for i, worker in enumerate(workers):
    #             earned = day_data.get(current, {}).get(worker, "")
    #             if earned != "":
    #                 totals[i] += earned
    #                 row.append(round(earned, 2))
    #             else:
    #                 row.append("")
    #         rows.append(row)
    #         current += timedelta(days=1)
    #
    #     total_row = ["Total"] + [round(t, 2) for t in totals]
    #     rows.append(total_row)
    #
    #     return rows


    def get_total_earned(self, worker, target_date):
        """
        Возвращает сумму, заработанную работником за установку материалов за указанный день.
        """
        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            worker=worker,
            date=target_date
        )

        total = 0
        for entry in entries:
            try:
                materials = entry.get_materials_used()
            except Exception as e:
                print(f"Ошибка при разборе JSON в WorkEntry(id={entry.id}): {e}")
                continue

            for name, quantity in materials.items():
                try:
                    material = Material.objects.get(name=name)
                    total += material.price * float(quantity)
                except Material.DoesNotExist:
                    print(f"Материал '{name}' не найден (WorkEntry id={entry.id})")
                except (ValueError, TypeError) as e:
                    print(f"Ошибка при обработке количества материала '{name}' в WorkEntry(id={entry.id}): {e}")

        return round(total, 2)

    def get_worker_day_earned(self, worker, target_date):
        """
        Возвращает KPI (заработано за установку - почасовая оплата) за день по объекту.
        """
        entries = WorkEntry.objects.filter(
            build_object=self.obj,
            worker=worker,
            date=target_date
        )

        total_hours = 0
        total_earned = 0

        for entry in entries:
            total_hours += float(entry.worked_hours)
            total_earned += self.get_total_earned(worker, target_date)  # предполагаем, что метод возвращает заработано за установку

        cost = worker.salary * total_hours
        kpi = total_earned - cost

        return round(kpi, 2)


    def get_start_row(self):
        """
        Возвращает строку, с которой начинать запись:
        - если в столбце A есть дата первого числа текущего месяца, возвращает её индекс + 1;
        - иначе использует базовый метод get_last_non_empty_row() + 1.
        """
        first_day_str = date.today().replace(day=1).strftime("%Y-%m-%d")

        try:
            # Читаем все ячейки столбца A
            column_a = f"{self.sheet_name}!A1:A"
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=column_a
            ).execute()
            values = result.get('values', [])
            for idx, row in enumerate(values):
                if row and row[0] == first_day_str:
                    return idx - 2 # строки в таблицах считаются с 1
        except Exception as e:
            print(f"Ошибка при чтении столбца A: {e}")

        # Если не нашли — используем базовый метод
        return self.get_last_non_empty_row() + 1

    def write_kpi_table(self):
        """
        Полностью перезаписывает таблицу KPI на листе.
        """
        today = date.today()
        first_day = today.replace(day=1)
        self.ensure_sheet_exists()

        # Проверка: есть ли вообще отчеты по объекту за месяц
        has_entries = WorkEntry.objects.filter(
            build_object=self.obj,
            date__gte=first_day,
            date__lte=today
        ).exists()

        if not has_entries:
            print(f"Пропуск: нет данных по объекту '{self.obj.name}' за {first_day.strftime('%B %Y')}")
            return

        rows = self.get_kpi_rows()
        print(f"Объект {self.obj.name} данные для записи {rows}")

        start_row = self.get_start_row()
        # Запись с первой строки
        self.update_data(f"A{start_row}", rows)

        # Применяем стили к основной таблице
        start_row += 2
        end_row = start_row + len(rows) - 3
        print(f"Форматиррвание, страт: {start_row} конец: {end_row} ")
        self.apply_styles(
            start_row=start_row,
            end_row=end_row,
            start_col=0,
            end_col=len(rows[2])
        )
        print(f"KPI-таблица обновлена: {len(rows) - 1} дней, {len(rows[2]) - 1} работников")


