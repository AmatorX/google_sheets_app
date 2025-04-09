from django.test import TestCase
from work_time_sheet import WorkTimeTable
from helpers.chunk import create_chunk_data


spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1Y2Wt6fLdDkjR2QZoEui0_2LtJCihzdTjOZXS4N5QNr4/edit?gid=0#gid=0'
worktime_table = WorkTimeTable(spreadsheet_url=spreadsheet_url)

rows = [
    {
        "per_hour": 25,
        "user_name": "John Doe",
        "work_hours": [8, 7, 6, 8, 9, 5, 0, 8, 8, 7, 6, 5, 8, 8]
    },
    {
        "per_hour": 20,
        "user_name": "Jane Smith",
        "work_hours": [7, 6, 8, 8, 5, 5, 0, 6, 7, 6, 5, 8, 8, 7]
    }
]

chunk_data = create_chunk_data()
worktime_table.create_table(date_chunk=chunk_data, rows=rows)