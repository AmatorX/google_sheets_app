# if __name__ == "__main__":
#     from sheets.work_time_sheet import WorkTimeTable
#     from utils.chunk import create_chunk_data
#
#     spreadsheet_url = 'https://docs.google.com/spreadsheets/d/13xzLus6iWfkMXGGziQj62T4lfgT3Lb_N55MCYE5Wz6c/edit?gid=0#gid=0'
#     worktime_table = WorkTimeTable(spreadsheet_url=spreadsheet_url)
#
#     # rows = [
#     #     {
#     #         "per_hour": 25,
#     #         "user_name": "John Doe",
#     #         "work_hours": [8, 7, 6, 8, 9, 5, 0, 8, 8, 7, 6, 5, 8, 11]
#     #     },
#     #     {
#     #         "per_hour": 20,
#     #         "user_name": "Jane Smith",
#     #         "work_hours": [7, 6, 8, 8, 5, 5, 0, 6, 7, 6, 5, 8, 8, 11]
#     #     }
#     # ]
#
#     rows = [
#         {'per_hour': 30, 'user_name': 'John Doe', 'work_hours': [8, 5, 8, 4, 4, 9, 6, 7, 7, 8, 8, 6, 9, 4]},
#         {'per_hour': 35, 'user_name': 'Jane Smith', 'work_hours': [6, 8, 5, 9, 6, 9, 6, 8, 4, 5, 8, 4, 5, 5]},
#         {'per_hour': 40, 'user_name': 'Michael Johnson', 'work_hours': [8, 6, 5, 4, 5, 6, 8, 5, 8, 6, 5, 5, 7, 5]},
#         {'per_hour': 40, 'user_name': 'Emily Davis', 'work_hours': [7, 4, 9, 4, 6, 7, 6, 7, 9, 9, 7, 7, 6, 7]},
#         {'per_hour': 30, 'user_name': 'James Brown', 'work_hours': [4, 5, 6, 7, 9, 8, 9, 4, 5, 4, 6, 6, 4, 6]},
#         {'per_hour': 25, 'user_name': 'Olivia Taylor', 'work_hours': [8, 8, 7, 8, 6, 6, 9, 5, 8, 7, 6, 9, 8, 6]},
#         {'per_hour': 28, 'user_name': 'Benjamin Lee', 'work_hours': [5, 7, 6, 5, 9, 8, 6, 5, 8, 7, 6, 5, 4, 9]},
#         {'per_hour': 38, 'user_name': 'Sophia Harris', 'work_hours': [6, 9, 8, 7, 9, 8, 8, 9, 7, 9, 8, 7, 8, 7]},
#         {'per_hour': 32, 'user_name': 'David Clark', 'work_hours': [7, 6, 7, 9, 8, 7, 6, 8, 5, 8, 7, 6, 9, 8]},
#         {'per_hour': 30, 'user_name': 'Charlotte Lewis', 'work_hours': [5, 7, 9, 7, 8, 6, 5, 6, 8, 5, 8, 9, 7, 7]},
#         {'per_hour': 45, 'user_name': 'Liam Walker', 'work_hours': [9, 8, 7, 5, 8, 9, 7, 6, 9, 6, 9, 7, 5, 8]},
#         {'per_hour': 32, 'user_name': 'Amelia Young', 'work_hours': [5, 6, 8, 9, 6, 8, 7, 6, 8, 7, 8, 7, 5, 6]},
#         {'per_hour': 38, 'user_name': 'Lucas Allen', 'work_hours': [9, 7, 5, 6, 6, 8, 9, 7, 8, 9, 8, 9, 6, 7]},
#         {'per_hour': 28, 'user_name': 'Harper King', 'work_hours': [6, 7, 5, 9, 9, 7, 6, 9, 6, 8, 9, 5, 7, 7]},
#         {'per_hour': 33, 'user_name': 'Ethan Scott', 'work_hours': [5, 6, 7, 8, 7, 5, 6, 7, 9, 6, 8, 7, 5, 8]},
#         {'per_hour': 36, 'user_name': 'Avery Wright', 'work_hours': [7, 5, 8, 9, 7, 8, 6, 9, 5, 6, 9, 6, 7, 9]},
#         {'per_hour': 31, 'user_name': 'Jack Adams', 'work_hours': [6, 7, 5, 8, 8, 9, 6, 8, 5, 7, 6, 5, 6, 7]},
#         {'per_hour': 34, 'user_name': 'Mason Green', 'work_hours': [9, 8, 6, 6, 7, 5, 7, 6, 8, 5, 9, 8, 9, 7]},
#         {'per_hour': 30, 'user_name': 'Ella Nelson', 'work_hours': [8, 7, 8, 9, 6, 6, 7, 6, 9, 5, 7, 9, 6, 6]},
#         {'per_hour': 42, 'user_name': 'Sebastian Carter', 'work_hours': [9, 9, 8, 8, 5, 7, 6, 7, 5, 6, 9, 7, 6, 6]},
#         {'per_hour': 33, 'user_name': 'Zoe Mitchell', 'work_hours': [8, 7, 5, 7, 8, 6, 9, 6, 8, 7, 7, 8, 9, 8]},
#         {'per_hour': 29, 'user_name': 'Daniel Perez', 'work_hours': [6, 9, 6, 7, 5, 8, 9, 7, 6, 9, 8, 5, 9, 5]},
#         {'per_hour': 40, 'user_name': 'Lily Roberts', 'work_hours': [9, 8, 5, 6, 7, 9, 8, 6, 8, 7, 5, 7, 6, 8]},
#         {'per_hour': 30, 'user_name': 'Henry Stewart', 'work_hours': [8, 6, 9, 5, 7, 5, 6, 9, 6, 8, 9, 7, 5, 9]},
#         {'per_hour': 36, 'user_name': 'Grace Morris', 'work_hours': [9, 6, 8, 7, 5, 9, 6, 8, 7, 6, 9, 9, 8, 8]},
#         {'per_hour': 35, 'user_name': 'Owen Rogers', 'work_hours': [6, 9, 8, 7, 9, 5, 6, 7, 5, 9, 7, 6, 9, 8]},
#         {'per_hour': 40, 'user_name': 'Samuel Reed', 'work_hours': [9, 7, 8, 6, 6, 5, 9, 6, 5, 8, 6, 5, 7, 7]},
#         {'per_hour': 32, 'user_name': 'Charlotte Peterson', 'work_hours': [8, 6, 8, 9, 7, 5, 9, 7, 6, 9, 5, 7, 8, 6]},
#         {'per_hour': 29, 'user_name': 'Joshua Cooper', 'work_hours': [9, 7, 5, 9, 8, 6, 6, 5, 7, 8, 7, 6, 8, 5]},
#         {'per_hour': 37, 'user_name': 'Isaac Bailey', 'work_hours': [6, 7, 9, 5, 9, 7, 9, 8, 7, 5, 9, 7, 6, 8]},
#         {'per_hour': 41, 'user_name': 'Victoria Sanchez', 'work_hours': [7, 8, 6, 5, 9, 9, 6, 7, 9, 5, 8, 6, 7, 9]},
#         {'per_hour': 28, 'user_name': 'Aidan Turner', 'work_hours': [6, 5, 8, 7, 6, 6, 9, 9, 7, 8, 7, 5, 9, 6]},
#         {'per_hour': 30, 'user_name': 'Nathan King', 'work_hours': [9, 5, 7, 6, 6, 7, 8, 8, 6, 5, 7, 6, 9, 7]},
#         {'per_hour': 32, 'user_name': 'Lila Hall', 'work_hours': [7, 6, 5, 7, 9, 8, 7, 9, 6, 9, 6, 7, 8, 8]},
#         {'per_hour': 35, 'user_name': 'Levi Adams', 'work_hours': [9, 8, 7, 5, 8, 9, 6, 9, 7, 8, 7, 6, 9, 5]},
#         {'per_hour': 28, 'user_name': 'Evelyn Young', 'work_hours': [6, 7, 8, 7, 5, 9, 6, 8, 9, 8, 7, 5, 6, 7]},
#         {'per_hour': 33, 'user_name': 'Luca Harris', 'work_hours': [5, 6, 8, 7, 9, 6, 8, 5, 6, 9, 7, 6, 5, 9]},
#         {'per_hour': 31, 'user_name': 'Mila Martinez', 'work_hours': [9, 6, 7, 9, 7, 5, 8, 6, 6, 9, 7, 8, 8, 6]},
#         {'per_hour': 34, 'user_name': 'Daniela Collins', 'work_hours': [8, 7, 9, 6, 7, 9, 8, 9, 5, 6, 8, 7, 6, 7]},
#         {'per_hour': 29, 'user_name': 'Gavin Allen', 'work_hours': [6, 7, 8, 9, 6, 7, 9, 8, 6, 7, 8, 5, 8, 9]},
#         {'per_hour': 40, 'user_name': 'Maddox Clark', 'work_hours': [9, 9, 5, 7, 6, 7, 6, 8, 7, 9, 6, 6, 9, 7]},
#         {'per_hour': 30, 'user_name': 'Maya Walker', 'work_hours': [5, 8, 7, 8, 6, 9, 8, 7, 6, 9, 7, 5, 8, 6]},
#         {'per_hour': 32, 'user_name': 'David Hill', 'work_hours': [9, 5, 6, 7, 9, 9, 8, 8, 9, 6, 8, 5, 7, 8]},
#         {'per_hour': 36, 'user_name': 'Aiden Young', 'work_hours': [6, 9, 7, 8, 5, 8, 9, 5, 7, 7, 8, 6, 7, 9]},
#         {'per_hour': 30, 'user_name': 'John Doe', 'work_hours': [8, 5, 8, 4, 4, 9, 6, 7, 7, 8, 8, 6, 9, 4]},
#         {'per_hour': 35, 'user_name': 'Jane Smith', 'work_hours': [6, 8, 5, 9, 6, 9, 6, 8, 4, 5, 8, 4, 5, 5]},
#         {'per_hour': 40, 'user_name': 'Michael Johnson', 'work_hours': [8, 6, 5, 4, 5, 6, 8, 5, 8, 6, 5, 5, 7, 5]},
#         {'per_hour': 40, 'user_name': 'Emily Davis', 'work_hours': [7, 4, 9, 4, 6, 7, 6, 7, 9, 9, 7, 7, 6, 7]},
#         {'per_hour': 30, 'user_name': 'James Brown', 'work_hours': [4, 5, 6, 7, 9, 8, 9, 4, 5, 4, 6, 6, 4, 6]},
#         {'per_hour': 25, 'user_name': 'Olivia Taylor', 'work_hours': [8, 8, 7, 8, 6, 6, 9, 5, 8, 7, 6, 9, 8, 6]},
#         {'per_hour': 28, 'user_name': 'Benjamin Lee', 'work_hours': [5, 7, 6, 5, 9, 8, 6, 5, 8, 7, 6, 5, 4, 9]},
#         {'per_hour': 38, 'user_name': 'Sophia Harris', 'work_hours': [6, 9, 8, 7, 9, 8, 8, 9, 7, 9, 8, 7, 8, 7]},
#         {'per_hour': 32, 'user_name': 'David Clark', 'work_hours': [7, 6, 7, 9, 8, 7, 6, 8, 5, 8, 7, 6, 9, 8]},
#         {'per_hour': 30, 'user_name': 'Charlotte Lewis', 'work_hours': [5, 7, 9, 7, 8, 6, 5, 6, 8, 5, 8, 9, 7, 7]},
#         {'per_hour': 45, 'user_name': 'Liam Walker', 'work_hours': [9, 8, 7, 5, 8, 9, 7, 6, 9, 6, 9, 7, 5, 8]},
#         {'per_hour': 32, 'user_name': 'Amelia Young', 'work_hours': [5, 6, 8, 9, 6, 8, 7, 6, 8, 7, 8, 7, 5, 6]},
#         {'per_hour': 38, 'user_name': 'Lucas Allen', 'work_hours': [9, 7, 5, 6, 6, 8, 9, 7, 8, 9, 8, 9, 6, 7]},
#         {'per_hour': 28, 'user_name': 'Harper King', 'work_hours': [6, 7, 5, 9, 9, 7, 6, 9, 6, 8, 9, 5, 7, 7]},
#         {'per_hour': 33, 'user_name': 'Ethan Scott', 'work_hours': [5, 6, 7, 8, 7, 5, 6, 7, 9, 6, 8, 7, 5, 8]},
#         {'per_hour': 36, 'user_name': 'Avery Wright', 'work_hours': [7, 5, 8, 9, 7, 8, 6, 9, 5, 6, 9, 6, 7, 9]},
#         {'per_hour': 31, 'user_name': 'Jack Adams', 'work_hours': [6, 7, 5, 8, 8, 9, 6, 8, 5, 7, 6, 5, 6, 7]},
#         {'per_hour': 34, 'user_name': 'Mason Green', 'work_hours': [9, 8, 6, 6, 7, 5, 7, 6, 8, 5, 9, 8, 9, 7]},
#         {'per_hour': 30, 'user_name': 'Ella Nelson', 'work_hours': [8, 7, 8, 9, 6, 6, 7, 6, 9, 5, 7, 9, 6, 6]},
#         {'per_hour': 42, 'user_name': 'Sebastian Carter', 'work_hours': [9, 9, 8, 8, 5, 7, 6, 7, 5, 6, 9, 7, 6, 6]},
#         {'per_hour': 33, 'user_name': 'Zoe Mitchell', 'work_hours': [8, 7, 5, 7, 8, 6, 9, 6, 8, 7, 7, 8, 9, 8]},
#         {'per_hour': 29, 'user_name': 'Daniel Perez', 'work_hours': [6, 9, 6, 7, 5, 8, 9, 7, 6, 9, 8, 5, 9, 5]},
#         {'per_hour': 40, 'user_name': 'Lily Roberts', 'work_hours': [9, 8, 5, 6, 7, 9, 8, 6, 8, 7, 5, 7, 6, 8]},
#         {'per_hour': 30, 'user_name': 'Henry Stewart', 'work_hours': [8, 6, 9, 5, 7, 5, 6, 9, 6, 8, 9, 7, 5, 9]},
#         {'per_hour': 36, 'user_name': 'Grace Morris', 'work_hours': [9, 6, 8, 7, 5, 9, 6, 8, 7, 6, 9, 9, 8, 8]},
#         {'per_hour': 35, 'user_name': 'Owen Rogers', 'work_hours': [6, 9, 8, 7, 9, 5, 6, 7, 5, 9, 7, 6, 9, 8]},
#         {'per_hour': 40, 'user_name': 'Samuel Reed', 'work_hours': [9, 7, 8, 6, 6, 5, 9, 6, 5, 8, 6, 5, 7, 7]},
#         {'per_hour': 32, 'user_name': 'Charlotte Peterson', 'work_hours': [8, 6, 8, 9, 7, 5, 9, 7, 6, 9, 5, 7, 8, 6]},
#         {'per_hour': 29, 'user_name': 'Joshua Cooper', 'work_hours': [9, 7, 5, 9, 8, 6, 6, 5, 7, 8, 7, 6, 8, 5]},
#         {'per_hour': 37, 'user_name': 'Isaac Bailey', 'work_hours': [6, 7, 9, 5, 9, 7, 9, 8, 7, 5, 9, 7, 6, 8]},
#         {'per_hour': 41, 'user_name': 'Victoria Sanchez', 'work_hours': [7, 8, 6, 5, 9, 9, 6, 7, 9, 5, 8, 6, 7, 9]},
#         {'per_hour': 28, 'user_name': 'Aidan Turner', 'work_hours': [6, 5, 8, 7, 6, 6, 9, 9, 7, 8, 7, 5, 9, 6]},
#         {'per_hour': 30, 'user_name': 'Nathan King', 'work_hours': [9, 5, 7, 6, 6, 7, 8, 8, 6, 5, 7, 6, 9, 7]},
#         {'per_hour': 32, 'user_name': 'Lila Hall', 'work_hours': [7, 6, 5, 7, 9, 8, 7, 9, 6, 9, 6, 7, 8, 8]},
#         {'per_hour': 35, 'user_name': 'Levi Adams', 'work_hours': [9, 8, 7, 5, 8, 9, 6, 9, 7, 8, 7, 6, 9, 5]},
#         {'per_hour': 28, 'user_name': 'Evelyn Young', 'work_hours': [6, 7, 8, 7, 5, 9, 6, 8, 9, 8, 7, 5, 6, 7]},
#         {'per_hour': 33, 'user_name': 'Luca Harris', 'work_hours': [5, 6, 8, 7, 9, 6, 8, 5, 6, 9, 7, 6, 5, 9]},
#         {'per_hour': 31, 'user_name': 'Mila Martinez', 'work_hours': [9, 6, 7, 9, 7, 5, 8, 6, 6, 9, 7, 8, 8, 6]},
#         {'per_hour': 34, 'user_name': 'Daniela Collins', 'work_hours': [8, 7, 9, 6, 7, 9, 8, 9, 5, 6, 8, 7, 6, 7]},
#         {'per_hour': 29, 'user_name': 'Gavin Allen', 'work_hours': [6, 7, 8, 9, 6, 7, 9, 8, 6, 7, 8, 5, 8, 9]},
#         {'per_hour': 40, 'user_name': 'Maddox Clark', 'work_hours': [9, 9, 5, 7, 6, 7, 6, 8, 7, 9, 6, 6, 9, 7]},
#         {'per_hour': 30, 'user_name': 'Maya Walker', 'work_hours': [5, 8, 7, 8, 6, 9, 8, 7, 6, 9, 7, 5, 8, 6]},
#         {'per_hour': 32, 'user_name': 'David Hill', 'work_hours': [9, 5, 6, 7, 9, 9, 8, 8, 9, 6, 8, 5, 7, 8]},
#         {'per_hour': 36, 'user_name': 'Aiden Young', 'work_hours': [6, 9, 7, 8, 5, 8, 9, 5, 7, 7, 8, 6, 7, 9]},
#         {'per_hour': 30, 'user_name': 'John Doe', 'work_hours': [8, 5, 8, 4, 4, 9, 6, 7, 7, 8, 8, 6, 9, 4]},
#         {'per_hour': 35, 'user_name': 'Jane Smith', 'work_hours': [6, 8, 5, 9, 6, 9, 6, 8, 4, 5, 8, 4, 5, 5]},
#         {'per_hour': 40, 'user_name': 'Michael Johnson', 'work_hours': [8, 6, 5, 4, 5, 6, 8, 5, 8, 6, 5, 5, 7, 5]},
#         {'per_hour': 40, 'user_name': 'Emily Davis', 'work_hours': [7, 4, 9, 4, 6, 7, 6, 7, 9, 9, 7, 7, 6, 7]},
#         {'per_hour': 30, 'user_name': 'James Brown', 'work_hours': [4, 5, 6, 7, 9, 8, 9, 4, 5, 4, 6, 6, 4, 6]},
#         {'per_hour': 25, 'user_name': 'Olivia Taylor', 'work_hours': [8, 8, 7, 8, 6, 6, 9, 5, 8, 7, 6, 9, 8, 6]},
#         {'per_hour': 28, 'user_name': 'Benjamin Lee', 'work_hours': [5, 7, 6, 5, 9, 8, 6, 5, 8, 7, 6, 5, 4, 9]},
#         {'per_hour': 38, 'user_name': 'Sophia Harris', 'work_hours': [6, 9, 8, 7, 9, 8, 8, 9, 7, 9, 8, 7, 8, 7]},
#         {'per_hour': 32, 'user_name': 'David Clark', 'work_hours': [7, 6, 7, 9, 8, 7, 6, 8, 5, 8, 7, 6, 9, 8]},
#         {'per_hour': 30, 'user_name': 'Charlotte Lewis', 'work_hours': [5, 7, 9, 7, 8, 6, 5, 6, 8, 5, 8, 9, 7, 7]},
#         {'per_hour': 45, 'user_name': 'Liam Walker', 'work_hours': [9, 8, 7, 5, 8, 9, 7, 6, 9, 6, 9, 7, 5, 8]},
#         {'per_hour': 32, 'user_name': 'Amelia Young', 'work_hours': [5, 6, 8, 9, 6, 8, 7, 6, 8, 7, 8, 7, 5, 6]},
#         {'per_hour': 38, 'user_name': 'Lucas Allen', 'work_hours': [9, 7, 5, 6, 6, 8, 9, 7, 8, 9, 8, 9, 6, 7]},
#         {'per_hour': 28, 'user_name': 'Harper King', 'work_hours': [6, 7, 5, 9, 9, 7, 6, 9, 6, 8, 9, 5, 7, 7]},
#         {'per_hour': 33, 'user_name': 'Ethan Scott', 'work_hours': [5, 6, 7, 8, 7, 5, 6, 7, 9, 6, 8, 7, 5, 8]},
#         {'per_hour': 36, 'user_name': 'Avery Wright', 'work_hours': [7, 5, 8, 9, 7, 8, 6, 9, 5, 6, 9, 6, 7, 9]},
#         {'per_hour': 31, 'user_name': 'Jack Adams', 'work_hours': [6, 7, 5, 8, 8, 9, 6, 8, 5, 7, 6, 5, 6, 7]},
#         {'per_hour': 34, 'user_name': 'Mason Green', 'work_hours': [9, 8, 6, 6, 7, 5, 7, 6, 8, 5, 9, 8, 9, 7]},
#         {'per_hour': 30, 'user_name': 'Ella Nelson', 'work_hours': [8, 7, 8, 9, 6, 6, 7, 6, 9, 5, 7, 9, 6, 6]},
#         {'per_hour': 42, 'user_name': 'Sebastian Carter', 'work_hours': [9, 9, 8, 8, 5, 7, 6, 7, 5, 6, 9, 7, 6, 6]},
#         {'per_hour': 33, 'user_name': 'Zoe Mitchell', 'work_hours': [8, 7, 5, 7, 8, 6, 9, 6, 8, 7, 7, 8, 9, 8]},
#         {'per_hour': 29, 'user_name': 'Daniel Perez', 'work_hours': [6, 9, 6, 7, 5, 8, 9, 7, 6, 9, 8, 5, 9, 5]},
#         {'per_hour': 40, 'user_name': 'Lily Roberts', 'work_hours': [9, 8, 5, 6, 7, 9, 8, 6, 8, 7, 5, 7, 6, 8]},
#         {'per_hour': 30, 'user_name': 'Henry Stewart', 'work_hours': [8, 6, 9, 5, 7, 5, 6, 9, 6, 8, 9, 7, 5, 9]},
#         {'per_hour': 36, 'user_name': 'Grace Morris', 'work_hours': [9, 6, 8, 7, 5, 9, 6, 8, 7, 6, 9, 9, 8, 8]},
#         {'per_hour': 35, 'user_name': 'Owen Rogers', 'work_hours': [6, 9, 8, 7, 9, 5, 6, 7, 5, 9, 7, 6, 9, 8]},
#         {'per_hour': 40, 'user_name': 'Samuel Reed', 'work_hours': [9, 7, 8, 6, 6, 5, 9, 6, 5, 8, 6, 5, 7, 7]},
#         {'per_hour': 32, 'user_name': 'Charlotte Peterson', 'work_hours': [8, 6, 8, 9, 7, 5, 9, 7, 6, 9, 5, 7, 8, 6]},
#         {'per_hour': 29, 'user_name': 'Joshua Cooper', 'work_hours': [9, 7, 5, 9, 8, 6, 6, 5, 7, 8, 7, 6, 8, 5]},
#         {'per_hour': 37, 'user_name': 'Isaac Bailey', 'work_hours': [6, 7, 9, 5, 9, 7, 9, 8, 7, 5, 9, 7, 6, 8]},
#         {'per_hour': 41, 'user_name': 'Victoria Sanchez', 'work_hours': [7, 8, 6, 5, 9, 9, 6, 7, 9, 5, 8, 6, 7, 9]},
#         {'per_hour': 28, 'user_name': 'Aidan Turner', 'work_hours': [6, 5, 8, 7, 6, 6, 9, 9, 7, 8, 7, 5, 9, 6]},
#         {'per_hour': 30, 'user_name': 'Nathan King', 'work_hours': [9, 5, 7, 6, 6, 7, 8, 8, 6, 5, 7, 6, 9, 7]},
#         {'per_hour': 32, 'user_name': 'Lila Hall', 'work_hours': [7, 6, 5, 7, 9, 8, 7, 9, 6, 9, 6, 7, 8, 8]},
#         {'per_hour': 35, 'user_name': 'Levi Adams', 'work_hours': [9, 8, 7, 5, 8, 9, 6, 9, 7, 8, 7, 6, 9, 5]},
#         {'per_hour': 28, 'user_name': 'Evelyn Young', 'work_hours': [6, 7, 8, 7, 5, 9, 6, 8, 9, 8, 7, 5, 6, 7]},
#         {'per_hour': 33, 'user_name': 'Luca Harris', 'work_hours': [5, 6, 8, 7, 9, 6, 8, 5, 6, 9, 7, 6, 5, 9]},
#         {'per_hour': 31, 'user_name': 'Mila Martinez', 'work_hours': [9, 6, 7, 9, 7, 5, 8, 6, 6, 9, 7, 8, 8, 6]},
#         {'per_hour': 34, 'user_name': 'Daniela Collins', 'work_hours': [8, 7, 9, 6, 7, 9, 8, 9, 5, 6, 8, 7, 6, 7]},
#         {'per_hour': 29, 'user_name': 'Gavin Allen', 'work_hours': [6, 7, 8, 9, 6, 7, 9, 8, 6, 7, 8, 5, 8, 9]},
#         {'per_hour': 40, 'user_name': 'Maddox Clark', 'work_hours': [9, 9, 5, 7, 6, 7, 6, 8, 7, 9, 6, 6, 9, 7]},
#         {'per_hour': 30, 'user_name': 'Maya Walker', 'work_hours': [5, 8, 7, 8, 6, 9, 8, 7, 6, 9, 7, 5, 8, 6]},
#         {'per_hour': 32, 'user_name': 'David Hill', 'work_hours': [9, 5, 6, 7, 9, 9, 8, 8, 9, 6, 8, 5, 7, 8]},
#         {'per_hour': 36, 'user_name': 'Aiden Young', 'work_hours': [6, 9, 7, 8, 5, 8, 9, 5, 7, 7, 8, 6, 7, 9]},
#         {'per_hour': 30, 'user_name': 'John Doe', 'work_hours': [8, 5, 8, 4, 4, 9, 6, 7, 7, 8, 8, 6, 9, 4]},
#         {'per_hour': 35, 'user_name': 'Jane Smith', 'work_hours': [6, 8, 5, 9, 6, 9, 6, 8, 4, 5, 8, 4, 5, 5]},
#         {'per_hour': 40, 'user_name': 'Michael Johnson', 'work_hours': [8, 6, 5, 4, 5, 6, 8, 5, 8, 6, 5, 5, 7, 5]},
#         {'per_hour': 40, 'user_name': 'Emily Davis', 'work_hours': [7, 4, 9, 4, 6, 7, 6, 7, 9, 9, 7, 7, 6, 7]},
#         {'per_hour': 30, 'user_name': 'James Brown', 'work_hours': [4, 5, 6, 7, 9, 8, 9, 4, 5, 4, 6, 6, 4, 6]},
#         {'per_hour': 25, 'user_name': 'Olivia Taylor', 'work_hours': [8, 8, 7, 8, 6, 6, 9, 5, 8, 7, 6, 9, 8, 6]},
#         {'per_hour': 28, 'user_name': 'Benjamin Lee', 'work_hours': [5, 7, 6, 5, 9, 8, 6, 5, 8, 7, 6, 5, 4, 9]},
#         {'per_hour': 38, 'user_name': 'Sophia Harris', 'work_hours': [6, 9, 8, 7, 9, 8, 8, 9, 7, 9, 8, 7, 8, 7]},
#         {'per_hour': 32, 'user_name': 'David Clark', 'work_hours': [7, 6, 7, 9, 8, 7, 6, 8, 5, 8, 7, 6, 9, 8]},
#         {'per_hour': 30, 'user_name': 'Charlotte Lewis', 'work_hours': [5, 7, 9, 7, 8, 6, 5, 6, 8, 5, 8, 9, 7, 7]},
#         {'per_hour': 45, 'user_name': 'Liam Walker', 'work_hours': [9, 8, 7, 5, 8, 9, 7, 6, 9, 6, 9, 7, 5, 8]},
#         {'per_hour': 32, 'user_name': 'Amelia Young', 'work_hours': [5, 6, 8, 9, 6, 8, 7, 6, 8, 7, 8, 7, 5, 6]},
#         {'per_hour': 38, 'user_name': 'Lucas Allen', 'work_hours': [9, 7, 5, 6, 6, 8, 9, 7, 8, 9, 8, 9, 6, 7]},
#         {'per_hour': 28, 'user_name': 'Harper King', 'work_hours': [6, 7, 5, 9, 9, 7, 6, 9, 6, 8, 9, 5, 7, 7]},
#         {'per_hour': 33, 'user_name': 'Ethan Scott', 'work_hours': [5, 6, 7, 8, 7, 5, 6, 7, 9, 6, 8, 7, 5, 8]},
#         {'per_hour': 36, 'user_name': 'Avery Wright', 'work_hours': [7, 5, 8, 9, 7, 8, 6, 9, 5, 6, 9, 6, 7, 9]},
#         {'per_hour': 31, 'user_name': 'Jack Adams', 'work_hours': [6, 7, 5, 8, 8, 9, 6, 8, 5, 7, 6, 5, 6, 7]},
#         {'per_hour': 34, 'user_name': 'Mason Green', 'work_hours': [9, 8, 6, 6, 7, 5, 7, 6, 8, 5, 9, 8, 9, 7]},
#         {'per_hour': 30, 'user_name': 'Ella Nelson', 'work_hours': [8, 7, 8, 9, 6, 6, 7, 6, 9, 5, 7, 9, 6, 6]},
#         {'per_hour': 42, 'user_name': 'Sebastian Carter', 'work_hours': [9, 9, 8, 8, 5, 7, 6, 7, 5, 6, 9, 7, 6, 6]},
#         {'per_hour': 33, 'user_name': 'Zoe Mitchell', 'work_hours': [8, 7, 5, 7, 8, 6, 9, 6, 8, 7, 7, 8, 9, 8]},
#         {'per_hour': 29, 'user_name': 'Daniel Perez', 'work_hours': [6, 9, 6, 7, 5, 8, 9, 7, 6, 9, 8, 5, 9, 5]},
#         {'per_hour': 40, 'user_name': 'Lily Roberts', 'work_hours': [9, 8, 5, 6, 7, 9, 8, 6, 8, 7, 5, 7, 6, 8]},
#         {'per_hour': 30, 'user_name': 'Henry Stewart', 'work_hours': [8, 6, 9, 5, 7, 5, 6, 9, 6, 8, 9, 7, 5, 9]},
#         {'per_hour': 36, 'user_name': 'Grace Morris', 'work_hours': [9, 6, 8, 7, 5, 9, 6, 8, 7, 6, 9, 9, 8, 8]},
#         {'per_hour': 35, 'user_name': 'Owen Rogers', 'work_hours': [6, 9, 8, 7, 9, 5, 6, 7, 5, 9, 7, 6, 9, 8]},
#         {'per_hour': 40, 'user_name': 'Samuel Reed', 'work_hours': [9, 7, 8, 6, 6, 5, 9, 6, 5, 8, 6, 5, 7, 7]},
#         {'per_hour': 32, 'user_name': 'Charlotte Peterson', 'work_hours': [8, 6, 8, 9, 7, 5, 9, 7, 6, 9, 5, 7, 8, 6]},
#         {'per_hour': 29, 'user_name': 'Joshua Cooper', 'work_hours': [9, 7, 5, 9, 8, 6, 6, 5, 7, 8, 7, 6, 8, 5]},
#         {'per_hour': 37, 'user_name': 'Isaac Bailey', 'work_hours': [6, 7, 9, 5, 9, 7, 9, 8, 7, 5, 9, 7, 6, 8]},
#         {'per_hour': 41, 'user_name': 'Victoria Sanchez', 'work_hours': [7, 8, 6, 5, 9, 9, 6, 7, 9, 5, 8, 6, 7, 9]},
#         {'per_hour': 28, 'user_name': 'Aidan Turner', 'work_hours': [6, 5, 8, 7, 6, 6, 9, 9, 7, 8, 7, 5, 9, 6]},
#         {'per_hour': 30, 'user_name': 'Nathan King', 'work_hours': [9, 5, 7, 6, 6, 7, 8, 8, 6, 5, 7, 6, 9, 7]},
#         {'per_hour': 32, 'user_name': 'Lila Hall', 'work_hours': [7, 6, 5, 7, 9, 8, 7, 9, 6, 9, 6, 7, 8, 8]},
#         {'per_hour': 35, 'user_name': 'Levi Adams', 'work_hours': [9, 8, 7, 5, 8, 9, 6, 9, 7, 8, 7, 6, 9, 5]},
#         {'per_hour': 28, 'user_name': 'Evelyn Young', 'work_hours': [6, 7, 8, 7, 5, 9, 6, 8, 9, 8, 7, 5, 6, 7]},
#         {'per_hour': 33, 'user_name': 'Luca Harris', 'work_hours': [5, 6, 8, 7, 9, 6, 8, 5, 6, 9, 7, 6, 5, 9]},
#         {'per_hour': 31, 'user_name': 'Mila Martinez', 'work_hours': [9, 6, 7, 9, 7, 5, 8, 6, 6, 9, 7, 8, 8, 6]},
#         {'per_hour': 34, 'user_name': 'Daniela Collins', 'work_hours': [8, 7, 9, 6, 7, 9, 8, 9, 5, 6, 8, 7, 6, 7]},
#         {'per_hour': 29, 'user_name': 'Gavin Allen', 'work_hours': [6, 7, 8, 9, 6, 7, 9, 8, 6, 7, 8, 5, 8, 9]},
#         {'per_hour': 40, 'user_name': 'Maddox Clark', 'work_hours': [9, 9, 5, 7, 6, 7, 6, 8, 7, 9, 6, 6, 9, 7]},
#         {'per_hour': 30, 'user_name': 'Maya Walker', 'work_hours': [5, 8, 7, 8, 6, 9, 8, 7, 6, 9, 7, 5, 8, 6]},
#         {'per_hour': 32, 'user_name': 'David Hill', 'work_hours': [9, 5, 6, 7, 9, 9, 8, 8, 9, 6, 8, 5, 7, 8]},
#         {'per_hour': 36, 'user_name': 'Aiden Young', 'work_hours': [6, 9, 7, 8, 5, 8, 9, 5, 7, 7, 8, 6, 7, 9]},
#     ]
#
#     chunk_data = create_chunk_data()
#     worktime_table.create_table(date_chunk=chunk_data, rows=rows)

# import os
# import django
#
# # Устанавливаем переменную окружения для настроек Django
# os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'tsa.settings')
#
# # Инициализируем Django, чтобы загрузить приложения
# django.setup()
# os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'tsa.settings')
# from sheets.work_time_sheet import WorkTimeTable
# from utils.chunk import create_chunk_data
#
# if __name__ == "__main__":
#     # URL для Google Spreadsheet
#     spreadsheet_url = 'https://docs.google.com/spreadsheets/d/13xzLus6iWfkMXGGziQj62T4lfgT3Lb_N55MCYE5Wz6c/edit?gid=0#gid=0'
#
#     # Создание объекта таблицы рабочих часов
#     worktime_table = WorkTimeTable(spreadsheet_url=spreadsheet_url)
#
#
#     # Генерация данных для текущего 14-дневного периода
#     chunk_data = create_chunk_data(full_list=False)
#     print(f'Chunk data: {chunk_data}')
#
#     worktime_table.ensure_sheet_exists()
#
#     # Удаляем старую таблицу, если есть
#     worktime_table.remove_existing_chunk_if_exists(chunk_data)
#
#     # Генерация строк для таблицы рабочих часов
#     chunk, rows_generated = worktime_table.generate_current_chunk_data_and_rows()
#     print(f'rows_generated: {rows_generated}')
#
#     # Вызов метода для создания таблицы в Google Sheets
#     worktime_table.create_table(chunk_data, rows_generated


import os
import django

# Устанавливаем переменную окружения для настроек Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'tsa.settings')
django.setup()

from sheets.work_time_sheet import WorkTimeTable
from utils.chunk import create_chunk_data
from buildings.models import BuildObject


if __name__ == "__main__":
    # Выбираем нужный BuildObject (по id)
    # build_object = BuildObject.objects.get(id=2)
    build_objects = BuildObject.objects.all()

    chunk_data = create_chunk_data(full_list=False)
    print(f'Chunk data: {chunk_data}')
    for build_object in build_objects:
        try:
            print(f'Работаем с объектом {build_object.name}')
            worktime_table = WorkTimeTable(obj=build_object)

            # Убедимся, что лист существует
            worktime_table.ensure_sheet_exists()

            # Удаляем старую таблицу, если есть
            worktime_table.remove_existing_chunk_if_exists(chunk_data)

            # Генерация строк для таблицы
            chunk, rows_generated = worktime_table.generate_current_chunk_data_and_rows()

            if rows_generated:
                worktime_table.create_table(chunk_data, rows_generated)
            else:
                print(f'Данных для записи в Work Time для объекта {build_object.name} нет.\n')
        except Exception as e:
            print(f'Ошибка при обработке объекта {build_object.name}: {e}\n')