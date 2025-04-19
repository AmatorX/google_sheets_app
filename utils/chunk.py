import datetime
import calendar


def create_chunk_data(full_list=False):
    """
    Формируем чанк для текущей недели начиная со второго воскресенья года.
    Если full_list=False, возвращается только текущий чанк.
    """
    year = datetime.datetime.now().year
    # Получение имен месяцев и аббревиатур дней недели один раз
    month_names = [calendar.month_name[month] for month in range(1, 13)]
    weekdays_abbr = list(calendar.day_abbr)
    # Создание списка всех дней в году с использованием списковых включений
    date_list = [
        (month_names[month - 1], weekdays_abbr[calendar.weekday(year, month, day)], str(day))
        for month in range(1, 13)
        for day in range(1, calendar.monthrange(year, month)[1] + 1)
    ]

    # Вычисление индекса второго воскресенья
    sundays = [i for i, (_, weekday, _) in enumerate(date_list) if weekday == 'Sun']
    second_sunday_index = sundays[1] if len(sundays) > 1 else None

    # Создание чанков
    chunked_date_list = [date_list[:second_sunday_index]] if second_sunday_index else []
    chunked_date_list.extend([date_list[i:i + 14] for i in range(second_sunday_index, len(date_list), 14)])

    if full_list:
        return chunked_date_list
    else:
        # Получаем текущую дату
        today = datetime.datetime.now()

        # Находим, в какой чанк попадает текущая дата
        for chunk in chunked_date_list:
            # Получаем первую дату в чанк
            chunk_start_date = datetime.datetime(year, month_names.index(chunk[0][0]) + 1, int(chunk[0][2]))

            # Проверяем, попадает ли сегодняшняя дата в этот чанк
            if chunk_start_date <= today < chunk_start_date + datetime.timedelta(days=14):
                return chunk

        return chunked_date_list[-1]  # В случае, если дата не попала в чанк, возвращаем последний чанк

