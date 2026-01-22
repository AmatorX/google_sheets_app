import datetime
import calendar


def create_chunk_data(full_list=False):
    """
    14-дневные чанки:
    - стандартно: воскресенье → суббота (14 дней)
    - первый чанк:
        * неполная первая неделя (с 1 января до первой субботы)
        * + следующая неделя до субботы
    - последний чанк может быть короче (обрезается 31 декабря)
    """
    today = datetime.date.today()
    year = today.year

    month_names = [calendar.month_name[m] for m in range(1, 13)]
    weekdays_abbr = list(calendar.day_abbr)

    def format_day(date_obj):
        return (
            month_names[date_obj.month - 1],
            weekdays_abbr[date_obj.weekday()],
            str(date_obj.day),
        )

    chunks = []

    start_date = datetime.date(year, 1, 1)
    end_of_year = datetime.date(year, 12, 31)

    # первый чанк (если 1 января не воскресенье)
    if start_date.weekday() != 6:  # 6 = Sunday
        chunk = []
        current = start_date
        saturdays_found = 0

        while current <= end_of_year:
            chunk.append(format_day(current))

            if current.weekday() == 5:  # Saturday
                saturdays_found += 1
                if saturdays_found == 2:
                    break

            current += datetime.timedelta(days=1)

        chunks.append(chunk)
        start_date = current + datetime.timedelta(days=1)

    # основные 14-дневные чанки (вс → сб)
    while start_date <= end_of_year:
        chunk = []

        for _ in range(14):
            if start_date > end_of_year:
                break
            chunk.append(format_day(start_date))
            start_date += datetime.timedelta(days=1)

        chunks.append(chunk)

    if full_list:
        return chunks

    # возвращаем текущий чанк
    for chunk in chunks:
        start = datetime.date(
            year,
            month_names.index(chunk[0][0]) + 1,
            int(chunk[0][2]),
        )
        end = start + datetime.timedelta(days=len(chunk))

        if start <= today < end:
            return chunk

    return chunks[-1]
