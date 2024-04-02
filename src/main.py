import datetime
import os
import queue
import shutil
import threading
import time
import uuid
from contextlib import contextmanager
from typing import Optional
import pyodbc
import flet as ft
from flet_core import icons, ButtonStyle, IconButton, colors, TimePickerEntryMode, MainAxisAlignment
import pandas as pd
import xlwings as xw

'''
class DataExtractor_Postgres:
    def __init__(self, host: str = 'localhost', port: int = 5432, database: str = 'app_database',
                 user: str = 'postgres', password: str = '123456789'):
        self.connection_string_sqlalchemy = f'postgresql+psycopg2://{user}:{password}@{host}:{port}/{database}'
        self.engine = create_engine(self.connection_string_sqlalchemy)

    # @timing_decorator
    def extract_data(self, sql_query: text) -> pd.DataFrame:
        df = pd.read_sql_query(sql_query, self.engine)
        return df
'''


def extract_data(server: str = 'BI-DEPT01', database: str = 'master', trusted_connection: str = 'yes', sql_query: str = None, timeout: int = 360) -> pd.DataFrame:
    connection_string = f"DRIVER=ODBC Driver 17 for SQL Server;SERVER={server};DATABASE={database};TRUSTED_CONNECTION={trusted_connection};"

    with pyodbc.connect(connection_string, timeout=timeout) as conn:
        cursor = conn.cursor()

        cursor.execute(sql_query)
        rows = cursor.fetchall()
        columns = [column[0] for column in cursor.description]

    # Закрытие соединения происходит автоматически при выходе из блока with
    df = pd.DataFrame.from_records(rows, columns=columns)

    return df


'''
def generate_unique_filename(base_path, base_filename):
    """Генерация уникального имени файла, добавляя дату и время."""
    timestamp = datetime.datetime.now().strftime("%Y_%m_%d_%H_%M")
    #print(timestamp)
    filename, extension = os.path.splitext(base_filename)
    return os.path.join(base_path, f"{filename}_{timestamp}{extension}")
'''


def csv_task(sql_query_path: str, csv_path: str, server: str, database: str):
    try:
        if not sql_query_path:
            return "Не выбран sql скрипт или путь для csv файла"
        # data_extractor = DataExtractor_Postgres()
        sql_query = open(sql_query_path).read()
        data = extract_data(server, database, sql_query=sql_query, timeout=1500)
        #print(data)
        if os.path.isdir(csv_path):
            # Если csv_path является директорией, генерируем уникальное имя файла
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H_%M")
            filename, extension = os.path.splitext(os.path.basename(sql_query_path))
            csv_path = os.path.join(csv_path, f"{filename}_{timestamp}.csv")

        data.to_csv(rf"{csv_path}", index=False, sep=";", )
        # Определяем диапазон данных
        return f"Файл обновлен: {csv_path}"
    except FileNotFoundError as e1:
        return f"Error: Файл не найден. {e1}"
    except UnicodeDecodeError:
        return f"Error: Ошибка декодирования файла {filename}."
    except PermissionError:
        return f"Error: Ошибка доступа к файлу {filename}. У вас нет прав на чтение или запись."
    except Exception as e:
        return f"Error: {e}"


def dependency(dependency_path: str, server: str, database: str):
    try:
        if not dependency_path:
            # Если dependency_path не существует, считаем зависимости выполненными
            return True
        # Проверка зависимостей в данном случае обновились ли таблицы
        sql_query = open(dependency_path).read()
        df = extract_data(server, database, sql_query=sql_query, timeout=60)
        # print(df)
        if df[df['DWH'] != 1].empty:
            # Зависимости выполнены
            #print(df)
            return True
        else:
            # print("Зависимости не выполнены")
            # Зависимости не выполнены
            return False
    except Exception as e:
        return f"Error: {e}"


def excel_task(excel_path: str, directory_path: str):
    try:
        # Открываем книгу
        # app = xw.App(visible=True)
        if not excel_path:
            return "Не выбран файл для обновления"
        wb = xw.Book(excel_path, read_only=False)
        # Задержка перед вызовом макроса
        # wb.app.visible = False
        time.sleep(1)
        # Обновляем файл
        wb.api.RefreshAll()
        time.sleep(2)
        # Вместо wb.app.quit() вызовите эту функцию
        # Устанавливаем таймер для закрытия приложения Excel через 15 минут
        # wb.app.kill()
        # Сохраняем изменения
        wb.save()
        # Задержка перед закрытием приложения Excel
        time.sleep(2)
        # Закрываем приложение Excel
        wb.app.quit()
        time.sleep(3)
        filename, extension = os.path.splitext(os.path.basename(excel_path))
        if (
            filename == "MinimalService_OTT_Internet_OldTariff_ATV" or filename == "REPORT_55") and directory_path != "":
            timestamp = datetime.datetime.now().strftime("%Y%m%d")
            new_filename = f"{filename}_{timestamp}{extension}"
            # Составляем новый путь для копии файла
            directory_path = os.path.join(directory_path, new_filename) if directory_path else new_filename

        if directory_path != "":
            shutil.copy2(excel_path, directory_path)
            return f"Файл обновлен: {filename} и скопирован в {directory_path}"

        return f"Файл обновлен: {filename}"
    except FileNotFoundError:
        return f"Error: Файл {filename} не найден."
    except PermissionError:
        return f"Error: Ошибка доступа к файлу {filename}. У вас нет прав на чтение или запись."
    except Exception as e:
        return f"Error: {e}"


def main(page: ft.Page):
    page.title = "PyFlow"
    # set the minimum width and height of the window.
    page.window_min_width = 960
    page.window_min_height = 480
    # Setting the theme of the page to light mode.
    page.theme_mode = "dark"
    # set the width and height of the window.
    page.window_width = 1600
    page.window_height = 720

    page.theme = ft.Theme(
        scrollbar_theme=ft.ScrollbarTheme(
            track_visibility=True,
            # track_border_color=ft.colors.BLUE,
            # thumb_visibility=True,
            thumb_color={
                ft.MaterialState.HOVERED: ft.colors.BLUE_GREY_500,
                ft.MaterialState.DEFAULT: ft.colors.BLUE_GREY_100,
            },
            thickness=10,
            radius=5,
            main_axis_margin=5,
            cross_axis_margin=10,
        )
    )

    def change_theme(e):
        page.theme_mode = "light" if page.theme_mode == "dark" else "dark"
        page.update()

    def execute_task(task):

        task.deactivate_button.disabled = True
        task.prog_ring.visible = True
        page.update()
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        # Проверяем зависимости
        dependencies_met = dependency(
            dependency_path=task.dependency_path.value,
            server=task.sql_dependency.server.value,
            database=task.sql_dependency.database.value,
        )

        if not dependencies_met:
            # Если зависимости не выполнены, переносим задачу в очередь через 1200 секунд
            thread = threading.Timer(1200, lambda: task_queue.put(task))
            thread.start()
            task.thread = thread
            tab_logs.content.controls.append(ft.Text(f"{current_time}: Tables for task {task.name.value} are not updated"))
            task.prog_ring.visible = False
            task.deactivate_button.disabled = False
            page.update()
            return
        # Продолжаем выполнение задачи
        if task.type == "default1":
            t = task.execute_func(
                sql_query_path=task.in_file_path.value,
                csv_path=task.out_file_path.value,
                server=task.sql_dialog.server.value,
                database=task.sql_dialog.database.value,
            )
        elif task.type == "default2":
            t = task.execute_func(
                excel_path=task.in_file_path.value,
                directory_path=task.out_file_path.value,
            )
        tab_logs.content.controls.append(ft.Text(f'{datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}: {t}'))
        task.prog_ring.visible = False
        task.deactivate_button.disabled = False
        if t.startswith("Файл обновлен"):
            task.last_update_time.value = f'{datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}'
        page.update()

        scheduled_time = datetime.datetime.strptime(task.schedule_time.value, "%H:%M:%S").time() #запланированное время
        selected_weekdays = task.segment_but.selected #выбраные дни недели
        current_week_day = datetime.datetime.now().weekday() #текущий день недели
        current_time = datetime.datetime.now().time() #текущее время
        
        # Находим ближайший выбранный день недели 
        if str(current_week_day) in selected_weekdays and len(selected_weekdays) == 1:
            print('Функция обновляется один раз в неделю')
            days_to_next_weekday = 7
        else:
            # Находим ближайший выбранный день недели
            next_weekdays = [int(day) for day in selected_weekdays if int(day) > current_week_day]
            if not next_weekdays:
                # Если нет выбранных дней на этой неделе, берем первый выбранный день на следующей неделе
                next_weekday = min(int(day) for day in selected_weekdays)
                days_to_next_weekday = (7 - current_week_day + next_weekday) % 7
            else:
                next_weekday = min(next_weekdays)
                days_to_next_weekday = next_weekday - current_week_day

        next_weekday_date = datetime.datetime.now() + datetime.timedelta(days=days_to_next_weekday)
        time_diff = datetime.datetime.combine(next_weekday_date.date(), scheduled_time) - datetime.datetime.combine(
            datetime.date.today(), current_time)
        print(time_diff)
        seconds_to_wait = max(time_diff.total_seconds(),0)
        print(seconds_to_wait)
            
        thread = threading.Timer(seconds_to_wait, lambda: task_queue.put(task))
        thread.start()
        task.thread = thread
        # print("close_execute_task func")


    def active_(task):
        scheduled_time = datetime.datetime.strptime(task.schedule_time.value, "%H:%M:%S").time()
        current_week_day = datetime.datetime.now().weekday()
        current_time = datetime.datetime.now().time()
        # Отключаем кнопку, чтобы избежать повторного запуска задачи
        task.active_button.disabled = True
        task.change_status()
        page.update()
        selected_weekdays = task.segment_but.selected
        if str(current_week_day) in selected_weekdays: 
            # Рассчитываем разницу во времени между текущим временем и запланированным временем
            time_diff = datetime.datetime.combine(datetime.date.today(), scheduled_time) - datetime.datetime.combine(
                datetime.date.today(), current_time)
            # Вычисляем количество секунд до запланированного времени, но не менее 0
            seconds_to_wait = max(time_diff.total_seconds(), 0)
            print(seconds_to_wait)
        else:
            # Находим ближайший выбранный день недели на следующей неделе
            next_weekdays = [int(day) for day in selected_weekdays if int(day) > current_week_day]
            if not next_weekdays:
                # Если нет выбранных дней на этой неделе, берем первый выбранный день на следующей неделе
                next_weekday = min(int(day) for day in selected_weekdays)
                days_to_next_weekday = (7 - current_week_day + next_weekday) % 7
            else:
                next_weekday = min(next_weekdays)
                days_to_next_weekday = next_weekday - current_week_day

            next_weekday_date = datetime.datetime.now() + datetime.timedelta(days=days_to_next_weekday)
            time_diff = datetime.datetime.combine(next_weekday_date.date(), scheduled_time) - datetime.datetime.combine(
                datetime.date.today(), current_time)
            print(time_diff)
            seconds_to_wait = max(time_diff.total_seconds(), 0)
            print(seconds_to_wait)
            
        thread = threading.Timer(seconds_to_wait, lambda: task_queue.put(task))
        thread.start()
        task.thread = thread
        # print("close active_ func")


    def deactivate_(task):
        if task.thread is not None:
            threading.Timer.cancel(task.thread)
            task.thread = None
            task.active_button.disabled = False
            task.change_status()
            page.update()
            # print("Task for Pipeline canceled")
        # else:
        # print(f"Задача {task.name} не активна или уже выполняется!")

    def create_task_queue():
        # Создаем экземпляр очереди для хранения задач
        task_queue = queue.Queue()
        # Создаем отдельный поток для обработки задач
        thread = threading.Thread(target=process_tasks, args=(task_queue,))
        # Запускаем поток
        thread.start()
        # Возвращаем созданную очередь
        return task_queue

    def process_tasks(task_queue):
        while True:
            # Получаем задачу из очереди
            task = task_queue.get()
            # print("задача в потоке")  # выводим что задача добавилась в поток
            # Выполняем задачу
            execute_task(task)
            # Сообщаем очереди, что задача выполнена
            task_queue.task_done()
            time.sleep(1)


    class segment_but(ft.SegmentedButton):
        def __init__(self, selected: set[str]):
            self.seg1 = ft.Segment(
                value="0",
                label=ft.Text("Mon"),
                # icon=ft.Icon(ft.icons.LOOKS_ONE),
            )
            self.seg2 = ft.Segment(
                value="1",
                label=ft.Text("Tue"),
                # icon=ft.Icon(ft.icons.LOOKS_ONE),
            )
            self.seg3 = ft.Segment(
                value="2",
                label=ft.Text("Wed"),
                # icon=ft.Icon(ft.icons.LOOKS_ONE),
            )
            self.seg4 = ft.Segment(
                value="3",
                label=ft.Text("Thu"),
                # icon=ft.Icon(ft.icons.LOOKS_ONE),
            )
            self.seg5 = ft.Segment(
                value="4",
                label=ft.Text("Fri"),
                # icon=ft.Icon(ft.icons.LOOKS_ONE),
            )
            self.seg6 = ft.Segment(
                value="5",
                label=ft.Text("Sat"),
                # icon=ft.Icon(ft.icons.LOOKS_ONE),
            )
            self.seg7 = ft.Segment(
                value="6",
                label=ft.Text("Sun"),
                # icon=ft.Icon(ft.icons.LOOKS_ONE),
            )
            super().__init__(selected_icon=ft.Icon(ft.icons.CIRCLE_SHARP),
                             scale=0.8,
                             width=590,
                             left=80,
                             selected=selected,
                             # on_change=lambda _: handle_change(self),
                             segments=[self.seg1, self.seg2,self.seg3,self.seg4,self.seg5,self.seg6,self.seg7],
                             style=ft.ButtonStyle(bgcolor={ft.MaterialState.SELECTED: ft.colors.INDIGO_300},
                                                  ),
                             allow_multiple_selection=True,
                             show_selected_icon=True,
                             )

    class progress_ring(ft.ProgressRing):
        def __init__(self):
            super().__init__(width=16, height=16, stroke_width=2)
            self.visible = False

        def change_visible(self):
            if self.visible == True:
                self.visible = False
            else:
                self.visible = True

    class sql_dialog(ft.AlertDialog):
        def __init__(self, task, save_func, server):
            self.task = task
            self.save_func = save_func
            self.server = ft.TextField(scale=0.85, width=200, value=server)
            self.database = ft.TextField(scale=0.85, width=200, value="master")
            self.FilePicker_sql = ft.FilePicker(
                on_result=lambda e: self.save_func(e))  # Файловый менеджер для выбора файлов
            page.overlay.append(self.FilePicker_sql)
            super().__init__(
                modal=True,
                content_padding=ft.margin.only(right=80, left=25),
                title=ft.Text("Заполните поля"),
                content=ft.Column(height=200,
                                  width=200,
                                  spacing=1,
                                  alignment=MainAxisAlignment.SPACE_AROUND,
                                  controls=[ft.Row(
                                      [ft.Text("Выберите файл"),
                                       ft.IconButton(
                                           scale=0.9,
                                           icon=ft.icons.UPLOAD_FILE,
                                           on_click=lambda _: self.FilePicker_sql.pick_files(
                                               allowed_extensions=["sql", "txt"]))
                                       ]),
                                      ft.Row(
                                          [ft.Text("SERVER", width=70), self.server],
                                      ),
                                      ft.Row(
                                          [ft.Text("DATABASE", width=70), self.database]),
                                  ]),
                actions=[
                    ft.TextButton("Confirm", on_click=lambda e: sql_dialog.confirm_dlg(self)),
                    ft.TextButton("Cancel", on_click=lambda e: sql_dialog.close_dlg(self)),
                ],
                actions_alignment=ft.MainAxisAlignment.END,
                # on_dismiss=lambda e: print("Modal dialog dismissed!"),
            )

        def close_dlg(self):
            self.open = False
            page.update()

        def confirm_dlg(self):
            # print(self.excel_sheet_name.value)
            # print(self.excel_table_name.value)
            self.open = False
            page.update()

        def open_dlg_modal(self):
            page.dialog = self
            self.open = True
            page.update()

    class info_dialog(ft.AlertDialog):
        def __init__(self, title: str):
            super().__init__(content=ft.Container(title, alignment=ft.alignment.center,height=35), 
                             content_padding=ft.margin.only(right=10, left=10, top=15),
                             )
        
        def open_info_dialog(self):
            page.dialog = self
            self.open = True
            page.update()

    class time_picker(ft.TimePicker):
        def __init__(self):
            super().__init__(confirm_text="Confirm",
                             error_invalid_text="Time out of range",
                             help_text="Pick your time slot",
                             # on_dismiss=dismissed,
                             time_picker_entry_mode=TimePickerEntryMode.INPUT_ONLY,
                             value=datetime.time(0, 0))
            self.task = None
            self.on_change = lambda _: self.apply_time(self.task)

        def apply_time(self, task):
            # print(f"{self.value}")
            if task:
                task.schedule_time.value = f"{self.value}"
            page.update()

        def pick_time(self, task):
            self.open = True
            # print(f"{self.value}")
            self.task = task
            self.update()

    class active_button(ft.IconButton):
        def __init__(self, row):
            super().__init__(icon=ft.icons.START, tooltip='Запустить')
            self.task = row
            self.on_click = lambda _: active_(self.task)

    class deactivate_button(ft.IconButton):
        def __init__(self, row):
            super().__init__(icon=ft.icons.STOP, tooltip='Остановить')
            self.task = row
            self.on_click = lambda _: deactivate_(self.task)

    class schedule_button(ft.ElevatedButton):
        def __init__(self, task):
            super().__init__(text="Schedule", icon=ft.icons.TIMER, scale=0.9, width=135)
            self.on_click = lambda _: time_picker.pick_time(task)

    class Text(ft.Text):
        def __init__(self, value):
            super().__init__(value=value, scale=1)


    class pipeline_row(ft.DataRow):
        def __init__(self,
                     name: str,
                     schedule_time: str,
                     last_update_time: str,
                     segment_but_selected: set[str] = {"1"},
                     ):
            # Инициализация атрибутов класса
            self.uuid = ft.Text(str(uuid.uuid4()))
            self.name = Text(name)  # имя нашей задачи
            self.status = Text("stopped")  # Статус задачи
            self.schedule_time = ft.Text(schedule_time)  # Время когда нужно выполнить задачу
            self.last_update_time = ft.Text(last_update_time)  # Время когда задача выполнялась последний раз
            self.prog_ring = progress_ring()  # Кольцо состояния загрузки(что задача в данный момент выполняется)
            self.segment_but = segment_but(selected=segment_but_selected)  # Кнопка для выбора дня недели для обновления задачи(pipeline-а)
            self.active_button = active_button(self)  # Кнопка для отправки задачи на выполнение
            self.deactivate_button = deactivate_button(self)  # Кнопка, что б убрать задачу с выполнения
            self.thread = None  # Ссылка на поток для нашей задачи
            # self.execute_func = None  # Ссылка на функцию для нашей задачи
            # Вызов конструктора родительского класса
            super().__init__(
                cells=[
                    ft.DataCell(ft.Row([ft.IconButton(
                        scale=0.8,
                        icon=ft.icons.DELETE,
                        on_click=lambda _: tab_all.table.remove_task(self)),
                        self.name])
                    ),
                    ft.DataCell(ft.Text("")),
                    ft.DataCell(ft.Text("")),
                    ft.DataCell(ft.Text("")),
                    ft.DataCell(self.status),
                    ft.DataCell(
                        ft.Stack(controls=[schedule_button(task=self), self.segment_but],expand=True,width=620)
                    ),
                    ft.DataCell(self.schedule_time),
                    ft.DataCell(self.last_update_time),
                    ft.DataCell(ft.Row(controls=[self.active_button, self.deactivate_button, self.prog_ring])),
                ])

        def change_status(self):
            if self.status.value == 'stopped':
                self.status.value = 'active'
            else:
                self.status.value = 'stopped'

    class task_sql_csv(pipeline_row):
        def __init__(self, type="", schedule_time: str = "00:00:00", last_update_time: str = "",
                     name: str = "task_sql_csv",
                     dependency_path: str = "",
                     in_file_path: str = "",
                     out_file_path: str = "",
                     server: str = "localhost",
                     segment_but_selected: set[str] = {"0"},
                     ):
            super().__init__(name=name, schedule_time=schedule_time, last_update_time=last_update_time, segment_but_selected=segment_but_selected)
            self.type = "default1"
            self.sql_dependency = sql_dialog(task=self, save_func=self.save_dependency_path, server=server)
            self.dependency_path = ft.Text(dependency_path)  # Переменная для хранения пути к sql файлу
            self.dependency = ft.Text(os.path.basename(dependency_path))
            self.execute_func = csv_task
            self.sql_dialog = sql_dialog(task=self, save_func=self.save_sql_path,
                                         server=server)  # Окно выбора excel файла, имя таблицы и имя листа
            self.in_file_path = ft.Text(in_file_path)  # Переменная для хранения пути к sql файлу
            #self.info_dialog = info_dialog(title=self.in_file_path)
            self.FilePicker_csv = ft.FilePicker(
                on_result=lambda e: self.save_csv_path(e))  # Файловый менеджер для выбора файлов
            self.FilePicker_directory = ft.FilePicker(
                on_result=lambda e: self.save_directory_path(e))
            self.out_file_path = ft.Text(out_file_path, width=180)  # Переменная для хранения пути к sql файлу
            page.overlay.append(self.FilePicker_csv)
            page.overlay.append(self.FilePicker_directory)
            self.cells[1] = ft.DataCell(ft.Row(controls=[
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.ADD_BOX,
                    on_click=lambda _: self.sql_dependency.open_dlg_modal()),
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.INFO,
                    on_click=lambda _: info_dialog(title=ft.Text(self.dependency_path.value)).open_info_dialog()),
                    self.dependency,
                ]
            ))
            self.cells[2] = ft.DataCell(ft.Row(controls=[
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.ADD_BOX,
                    on_click=lambda _: self.sql_dialog.open_dlg_modal()),
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.INFO,
                    on_click=lambda _: info_dialog(title=ft.Text(self.in_file_path.value)).open_info_dialog()),
                ]
            ))
            self.cells[3] = ft.DataCell(ft.Row(controls=[
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.UPLOAD_FILE,
                    on_click=lambda _: self.FilePicker_csv.pick_files(
                        allowed_extensions=["csv"])),
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.DRIVE_FOLDER_UPLOAD,
                    on_click=lambda _: self.FilePicker_directory.get_directory_path()),
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.INFO,
                    on_click=lambda _: info_dialog(title=ft.Text(self.out_file_path.value)).open_info_dialog()),
            ]))

        def save_csv_path(self, e: ft.FilePickerResultEvent):
            self.out_file_path.value = (
                ", ".join(map(lambda f: f.path, e.files)) if e.files else ""
            )
            page.update()

        def save_directory_path(self, e: ft.FilePickerResultEvent):
            self.out_file_path.value = e.path if e.path else ""
            page.update()

        def save_sql_path(self, e: ft.FilePickerResultEvent):
            self.in_file_path.value = (
                ", ".join(map(lambda f: f.path, e.files)) if e.files else ""
            )
            self.name.value = (
                ", ".join(map(lambda f: f.name, e.files)) if e.files else self.name.value
            )
            page.update()

        def save_dependency_path(self, e: ft.FilePickerResultEvent):
            self.dependency_path.value = (
                ", ".join(map(lambda f: f.path, e.files)) if e.files else ""
            )
            self.dependency.value = (
                ", ".join(map(lambda f: f.name, e.files)) if e.files else self.dependency.value
            )
            page.update()

    class task_excel(pipeline_row):
        def __init__(self, type="", schedule_time: str = "00:00:00", last_update_time: str = "",
                     name: str = "task_excel",
                     dependency_path: str = "",
                     in_file_path: str = "",
                     out_file_path: str = "",
                     server: str = "",
                     segment_but_selected: set[str] = {"0"},
                     ):
            super().__init__(name=name, schedule_time=schedule_time, last_update_time=last_update_time, segment_but_selected=segment_but_selected)
            self.type = "default2"
            self.execute_func = excel_task
            self.sql_dependency = sql_dialog(
                task=self, save_func=self.save_dependency_path,
                server=server)  # Окно выбора sql файла, в данном случае запрос будет проверять наличие актуальных данных
            self.dependency_path = ft.Text(dependency_path)  # Переменная для хранения пути к sql файлу
            self.dependency = ft.Text(os.path.basename(dependency_path))
            self.in_file_path = ft.Text(in_file_path)
            self.out_file_path = ft.Text(out_file_path)
            self.FilePicker_excel = ft.FilePicker(
                on_result=lambda e: self.save_excel_path(e))  # Файловый менеджер для выбора файлов
            self.FilePicker_directory = ft.FilePicker(
                on_result=lambda e: self.save_directory_path(e))
            page.overlay.append(self.FilePicker_excel)
            page.overlay.append(self.FilePicker_directory)
            self.cells[1] = ft.DataCell(ft.Row(controls=[
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.ADD_BOX,
                    on_click=lambda _: self.sql_dependency.open_dlg_modal()),
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.INFO,
                    on_click=lambda _: info_dialog(title=ft.Text(self.dependency_path.value)).open_info_dialog()),
                    self.dependency,
                ]
            ))
            self.cells[2] = ft.DataCell(ft.Row(controls=[
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.UPLOAD_FILE,
                    on_click=lambda _: self.FilePicker_excel.pick_files(
                        allowed_extensions=["xlsm", "xlsb", "xls", "xltm", "xla", "xlsx", "xltx"])),
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.INFO,
                    on_click=lambda _: info_dialog(title=ft.Text(self.in_file_path.value)).open_info_dialog()),
                ]
            ))
            self.cells[3] = ft.DataCell(ft.Row(controls=[
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.DRIVE_FOLDER_UPLOAD,
                    on_click=lambda _: self.FilePicker_directory.get_directory_path()),
                ft.IconButton(
                    scale=0.9,
                    icon=ft.icons.INFO,
                    on_click=lambda _: info_dialog(title=ft.Text(self.out_file_path.value)).open_info_dialog()),
            ]))

        def save_excel_path(self, e: ft.FilePickerResultEvent):
            self.in_file_path.value = (
                ", ".join(map(lambda f: f.path, e.files)) if e.files else ""
            )
            self.name.value = (
                ", ".join(map(lambda f: f.name, e.files)) if e.files else self.name.value
            )
            page.update()

        def save_dependency_path(self, e: ft.FilePickerResultEvent):
            self.dependency_path.value = (
                ", ".join(map(lambda f: f.path, e.files)) if e.files else ""
            )
            self.dependency.value = (
                ", ".join(map(lambda f: f.name, e.files)) if e.files else self.dependency.value
            )
            

        def save_directory_path(self, e: ft.FilePickerResultEvent):
            self.out_file_path.value = e.path if e.path else ""
            page.update()

    class table(ft.DataTable):
        def __init__(self, self_tab, rows: list[task_sql_csv | task_excel]):
            super().__init__(heading_row_color=ft.colors.BLACK12,
                             data_row_max_height=65,
                             data_row_min_height=50,
                             heading_row_height=40,
                             column_spacing=25,
                             columns=[
                                 ft.DataColumn(ft.Text(value="Task")),
                                 ft.DataColumn(ft.Text(value="dependencies")),
                                 ft.DataColumn(ft.Text(value="input file")),
                                 ft.DataColumn(ft.Text(value="output file/path")),
                                 ft.DataColumn(ft.Text(value="Status")),
                                 ft.DataColumn(ft.Text("Schedule")),
                                 ft.DataColumn(ft.Text("Time update")),
                                 ft.DataColumn(ft.Text("Last Run")),
                                 ft.DataColumn(ft.Text("Actions")),
                             ])
            self.rows = rows
            self.self_tab = self_tab

        def create_task(self, task=None):
            """Метод для создания задачи в таблице."""
            if task is None:
                task = task_sql_csv()  # Создаем новую задачу, если не передана
            if task not in self.rows:
                self.rows.append(task)
                self.self_tab.update_tab()
                page.show_snack_bar(
                    ft.SnackBar(ft.Text("Task was added successfully!"), open=True, duration=1500)
                )
            else:
                page.show_snack_bar(
                    ft.SnackBar(ft.Text("Task already in the table."), open=True, duration=1500)
                )
            page.update()

        def add_task(self, task: task_sql_csv):
            """ Метод для добавления задачи в таблицу. """
            if task not in self.rows:
                self.rows.append(task)
            # else:
            # print("Task already in the table.")
            page.update()

        def remove_task(self, task: task_sql_csv | task_excel):
            """ Метод для удаления задачи из таблицы. """
            if task in self.rows:
                self.rows.remove(task)
                self.self_tab.update_tab()
            # else:
            # print("Task not found in the table.")
            page.update()


    class tab(ft.Tab):
        def __init__(self, tab_content: str, rows: list = []):
            # self.text = str(len(content.rows))
            # self.tab_content = ft.Badge(content=ft.Text(value=tab_content), text=str(len(rows)))
            self.table = table(rows=rows, self_tab=self)
            super().__init__(content=ft.Row(vertical_alignment=ft.CrossAxisAlignment.START,
                                            controls=[ft.Container(content=ft.Column(
                                                controls=[self.table],
                                                scroll="auto"),
                                                bgcolor=ft.colors.WHITE12,
                                                margin=ft.margin.only(bottom=30)
                                            )
                                            ], scroll="always"
                                            ),
                             tab_content=ft.Badge(content=ft.Text(value=tab_content), text=str(len(self.table.rows)))
                             )

        def update_tab(self):
            self.tab_content.text = str(len(self.table.rows)) # обновляем счетчик кол-ва задач 

    def save_tasks(tab):
        page.client_storage.clear()
        for task in tab.table.rows:
            if page.client_storage.contains_key(f"{task.uuid.value}"):
                page.client_storage.remove(f"{task.uuid.value}")
                
            page.client_storage.set(f"pyflow.{task.uuid.value}", {"name": task.name.value,
                                                           "type": task.type,
                                                           "dependency_path": task.dependency_path.value,
                                                           "in_file_path": task.in_file_path.value,
                                                           "out_file_path": task.out_file_path.value,
                                                           "schedule_time": task.schedule_time.value,
                                                           "last_update_time": task.last_update_time.value,
                                                           "server": task.sql_dependency.server.value,
                                                           "segment_but_selected": [i for i in task.segment_but.selected]
                                                           })
        page.show_snack_bar(
            ft.SnackBar(ft.Text("All tasks saved"), open=True, duration=1500)
        )

    def load_tasks():
        tasks = page.client_storage.get_keys("pyflow.")
        # print(tasks)

        for task in tasks:
            data = page.client_storage.get(task)

            # Определяем, какой тип задачи
            task_type = data.get("type", "default1")  # Если атрибут "type" не указан, используем значение по умолчанию

            # В зависимости от типа вызываем соответствующий конструктор
            if task_type == "default1":
                # print(data)
                task_object = task_sql_csv(**data)
            elif task_type == "default2":
                # print(data)
                task_object = task_excel(**data)
            else:
                # print(f"Неизвестный тип задачи: {task_type}")
                continue
            if task_object not in tab_all.table.rows:
                tab_all.table.create_task(task_object)
            # else:
            # print(f"Об'єкт с таким именем уже существует")

    task_queue = create_task_queue()

    add_task_1 = ft.ElevatedButton(icon=ft.icons.ADD_CIRCLE, text="update excel",
                                   on_click=lambda _: tab_all.table.create_task(task=task_excel()))

    add_task_2 = ft.ElevatedButton(icon=ft.icons.ADD_CIRCLE, text="load sql to csv",
                                   on_click=lambda _: tab_all.table.create_task(task=task_sql_csv()))

    save_tasks_button = ft.IconButton(icon=ft.icons.SAVE, on_click=lambda _: save_tasks(tab_all))

    time_picker = time_picker()  # Часы, что б выбрать значение для задачи schedule_time
    page.overlay.append(time_picker)
    tab_all = tab(tab_content="All  ")

    # page.client_storage.clear()

    tab_logs = ft.Tab(text="Logs", content=ft.ListView(expand=1, spacing=10, padding=20, ))

    main_tab = ft.Tabs(
        selected_index=1,
        animation_duration=350,
        tabs=[
            tab_all,
            tab_logs
        ],
        expand=1
    )
    theme_icon_button = IconButton(
        icons.DARK_MODE,
        selected_icon=icons.LIGHT_MODE,
        icon_color=colors.BLACK,
        icon_size=25,
        tooltip="change theme",
        on_click=change_theme,
        style=ButtonStyle(
            color={"": colors.BLACK, "selected": colors.WHITE},
        ),
    )

    page.add(save_tasks_button, main_tab, ft.Row([theme_icon_button, add_task_1, add_task_2]))
    load_tasks()  # загрузить все задачи которые были сохраненны ранее


if __name__ == "__main__":
    ft.app(target=main)
