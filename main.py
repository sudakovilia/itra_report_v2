from datetime import datetime, timedelta
from threading import Thread
from tkcalendar import DateEntry
from tkinter import E, filedialog as fd
from tkinter import messagebox as mb
from tkinter import ttk
import json
import pandas as pd
import tkinter as tk
import xlsxwriter


class View(tk.Tk):

    def __init__(self, master=None) -> None:
        super().__init__()

        self.title('ITRA reports')
        self.minsize(400, 300)

        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(side='top', fill='both', expand=True, padx=10, pady=10)

        current_row = 0
        employee_label = ttk.Label(self.main_frame, text='Список сотрудников')
        employee_label.grid(row=current_row, column=0, sticky='we')
        self.employee_file_path = tk.StringVar()
        employee_button = ttk.Button(self.main_frame, text='Выбрать файл', command=lambda:self.select_file(self.employee_file_path))
        employee_button.grid(row=current_row, column=1, sticky='we')

        current_row += 1
        staffing_label = ttk.Label(self.main_frame, text='Стаффинг')
        staffing_label.grid(row=current_row, column=0, sticky='we')
        self.staffing_file_path = tk.StringVar()
        staffing_button = ttk.Button(self.main_frame, text='Выбрать файл', command=lambda:self.select_file(self.staffing_file_path))
        staffing_button.grid(row=current_row, column=1, sticky='we')

        current_row += 1
        charging_label = ttk.Label(self.main_frame, text='Чарджинг')
        charging_label.grid(row=current_row, column=0, sticky='w')
        self.charging_file_path = tk.StringVar()
        charging_button = ttk.Button(self.main_frame, text='Выбрать файл', command=lambda:self.select_file(self.charging_file_path))
        charging_button.grid(row=current_row, column=1, sticky='we')

        current_row += 1
        date_from_label = ttk.Label(self.main_frame, text='Дата ОТ')
        date_from_label.grid(row=current_row, column=0, sticky='w')
        self.date_from_str = tk.StringVar()
        date_from_date_entry = DateEntry(self.main_frame, select_mode='day', textvariable=self.date_from_str)
        date_from_date_entry.grid(row=current_row, column=1, sticky='we')
        
        current_row += 1
        date_to_label = ttk.Label(self.main_frame, text='Дата ДО')
        date_to_label.grid(row=current_row, column=0, sticky='w')
        self.date_to_str = tk.StringVar()
        date_to_date_entry = DateEntry(self.main_frame, select_mode='day', textvariable=self.date_to_str)
        date_to_date_entry.grid(row=current_row, column=1, sticky='we')

        current_row += 1
        report_label = ttk.Label(self.main_frame, text='Отчет')
        report_label.grid(row=current_row, column=0, sticky='w')
        combo_values = [
            'Стаффинг формальный',
            'Стаффинг внутренний мониторинг',
            'Сверка стаффинг-чарджинг'
        ]
        self.report_combo = ttk.Combobox(self.main_frame, values=combo_values)
        self.report_combo.current(0)
        self.report_combo.grid(row=current_row, column=1, sticky='we')

        current_row += 1
        self.generate_report_button = ttk.Button(self.main_frame, text='Сформировать отчет', width=40, command=self.generate_report)
        self.generate_report_button.grid(row=current_row, column=0, columnspan=2)
        self.generate_report_button['state'] = 'disabled'


        for row_n in range(current_row+1):
            self.main_frame.rowconfigure(row_n, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=4)

    def select_file(self, file_path_var: tk.StringVar):
        filetypes = (
            ('Excel files', '*.xlsx'),
            ('All files', '*.*')
        )

        file_path = fd.askopenfilename(
            title='Выберите файл',
            initialdir='.',
            filetypes=filetypes)

        if file_path != '':
            file_path_var.set(file_path)
            if self.check_all_files_selected():
                self.generate_report_button['state'] = 'normal'
    
    def check_all_files_selected(self):
        file_paths = (
            self.staffing_file_path.get(),
            self.charging_file_path.get(),
            self.employee_file_path.get()
        )
        return False if '' in file_paths else True

    def check_dates_valid(self):
        date_from = datetime.strptime(self.date_from_str.get(), '%m/%d/%y').date()
        date_to = datetime.strptime(self.date_to_str.get(), '%m/%d/%y').date()
        report_selected = self.report_combo.current()

        if date_from >= date_to:
            return False

        if date_from.weekday() != 0 or date_to.weekday() != 4:
            return False
                
        if report_selected == 2 and (date_to - date_from).days != 4:
            return False

        return True

    def generate_report(self):
        if not(self.check_dates_valid()):
            msg = 'Проверьте что:\n1. Дата ОТ < Даты ДО\n2. Дата ОТ - понедельник\n3. Дата ДО - пятница\n4. Период сверки - 1 неделя'
            mb.showerror(title='Ошибка', message=msg)
            return 1
        
        self.generate_report_button['state'] = 'disabled'
        self.main_frame.config(cursor='wait')
        view_context = {
            'employee_file_path': self.employee_file_path.get(),
            'staffing_file_path': self.staffing_file_path.get(),
            'charging_file_path': self.charging_file_path.get(),
            'date_from': datetime.strptime(self.date_from_str.get(), '%m/%d/%y').date(),
            'date_to': datetime.strptime(self.date_to_str.get(), '%m/%d/%y').date(),
            'report_combo': self.report_combo.current()
        }
        thread = ReportGenerationThread(view_context)
        thread.start()
        self.monitor(thread)


    def monitor(self, thread):
        if thread.is_alive():
            self.after(100, lambda: self.monitor(thread))
        else:
            self.generate_report_button['state'] = 'normal'
            self.main_frame.config(cursor='')
            if thread.result_msg['status'] == 'ok':
                mb.showinfo(title='Готово', message=thread.result_msg['message'])
            else:
                mb.showerror(title='Ошибка', message=thread.result_msg['message'])
        
    def main(self):
        self.mainloop()

class ReportGenerationThread(Thread):

    def __init__(self, view_context: dict):
        super().__init__()
        self.view_context = view_context
        self.result_msg = {'status': 'error', 'message': 'Report generation error'}

    def run(self):
        if self.view_context['report_combo'] == 2:
            generator = StaffingVsChargingReportGenerator(self.view_context)
        else:
            generator = StaffingReportGenerator(self.view_context)

        self.result_msg = generator.result_msg

class StaffingReportGenerator:

    def __init__(self, view_context: dict) -> None:
        try:
            staffing_loader = StaffingDataLoader(
                view_context['staffing_file_path'],
                view_context['date_from'],
                view_context['date_to']
            )
        except Exception as e:
            self.result_msg = {'status': 'error', 'message': 'Ошибка загрузки стаффинга'}
            return None
        self.staffing_df = staffing_loader.get_df()
        self.week_cols = staffing_loader.get_week_cols()
        
        try:
            self.staffing_cell_generator = StaffingReportCellGenerator(
                staffing_loader.get_df(),
                1 if view_context['report_combo'] == 0 else 2
            )
        except Exception as e:
            self.result_msg = {'status': 'error', 'message': 'Ошибка загрузки файла formats.json'}
            return None
        
        try:
            employee_data_loader = EmployeeDataLoader(view_context['employee_file_path'])
        except Exception as e:
            self.result_msg = {'status': 'error', 'message': 'Ошибка загрузки списка сотрудников'}
            return None

        self.employee_df = employee_data_loader.get_employee_df()
        self.employee_list = employee_data_loader.get_employee_list()
        
        week_name = view_context['date_from'].strftime('%d-%m-%Y')
        self.save_path = f'Staffing_ITRA_byPerson-w-{week_name}.xlsx'
        
        self.set_up_excel_workbook()
        self.set_formats()
        self.print_staff_info()
        self.print_week_cols()
        self.print_report_data()
        self.workbook.close()
        self.result_msg = {'status': 'ok', 'message': f'Отчет сохранен в файл {self.save_path}'}


    def set_up_excel_workbook(self):
        workbook = xlsxwriter.Workbook(self.save_path)
        worksheet = workbook.add_worksheet('Staffing_report')
        worksheet.freeze_panes(1, 2)
        worksheet.set_zoom(50)
        worksheet.set_column(0, 0, 25)  # ширина колонки с именанми
        worksheet.set_column(1, 1, 10)  # ширина колонки с грейдами
        worksheet.set_column(2, len(self.week_cols) + 1, 50)  # ширина колонок с инфой
        for n in range(1, len(self.employee_list) + 1):
            worksheet.set_row(n, 150)
        
        self.workbook = workbook
        self.worksheet = worksheet

    def set_formats(self):
        self.header_format = self.workbook.add_format({
            'bold': True,
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#a9a9a9',
            'border': 1
        })
        self.spec_format = self.workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })


    def print_staff_info(self):
        self.worksheet.write(0, 0, 'Specialist', self.header_format)
        self.worksheet.write(0, 1, 'Grade', self.header_format)

        for n, employee in enumerate(self.employee_list):
            self.worksheet.write(n + 1, 0, employee[1], self.spec_format)
            self.worksheet.write(n + 1, 1, employee[2], self.spec_format)

    def print_week_cols(self):
        for n, week in enumerate(self.week_cols):
            week_label = week.strftime('%d %B %Y')
            self.worksheet.write(0, n + 2, week_label, self.header_format)

    def print_report_data(self):
        for employee_n, employee in enumerate(self.employee_list):
            for week_n, week in enumerate(self.week_cols):
                text = self.staffing_cell_generator.get_cell_text(employee[0], week)
                format = self.workbook.add_format(self.staffing_cell_generator.get_cell_format(employee[0], week))
                self.worksheet.write(employee_n + 1, week_n + 2, text, format)

class StaffingVsChargingReportGenerator:
    
    def __init__(self, view_context: dict) -> None:
        date_from=view_context['date_from']
        date_to=view_context['date_to']

        try:
            employee = EmployeeDataLoader(view_context['employee_file_path'])
        except Exception as e:
            self.result_msg = {'status': 'error', 'message': 'Ошибка загрузки списка сотрудников'}
            return None
        employee_df = employee.get_employee_df()

        try:
            staffing = StaffingDataLoader(view_context['staffing_file_path'])
        except Exception as e:
            self.result_msg = {'status': 'error', 'message': 'Ошибка загрузки стаффинга'}
            return None
        staffing_total = staffing.get_total_df(date_from, date_to)

        try:
            staffing_cell_generator = StaffingReportCellGenerator(staffing.get_df(), 1)
        except:
            self.result_msg = {'status': 'error', 'message': 'Ошибка загрузки файла formats.json'}
            return None

        try:
            charging = ChargingDataLoader(view_context['charging_file_path'])
        except Exception as e:
            self.result_msg = {'status': 'error', 'message': 'Ошибка загрузки чарджинга'}
            return None

        charging_total = charging.get_total_df(date_from, date_to)

        report = pd.merge(employee_df, staffing_total, how='left', left_index=True, right_index=True, sort=False)
        report = pd.merge(report, charging_total, how='left', left_index=True, right_index=True, sort=False)
        report.fillna(value={'Charged on client codes': 0, 'Staffing (total)': 0}, inplace=True)
        report['Charging - Staffing'] = report['Charged on client codes'] - report['Staffing (total)']
        report.sort_values(by=['grade_order', 'Name'], inplace=True)
        report['Project manager'] = ''

        # вывод сверки в файл
        try:
            formatter = CellFormatter()
        except Exception as e:
            self.result_msg = {'status': 'error', 'message': f'Ошибка загрузки файла formats.json'}
            return None
            
        output_file_name = 'Staffing vs Charging_week {}-{}.xlsx'.format(date_from.strftime("%d.%m"),
                                                                         date_to.strftime("%d.%m.%Y"))

        workbook = xlsxwriter.Workbook(output_file_name)
        worksheet = workbook.add_worksheet('Report')
        worksheet.set_zoom(65)
        worksheet.set_row(0, 66)
        worksheet.set_row(1, 46)

        for row in range(2, len(report.index) + 2):
            worksheet.set_row(row, 104)

        worksheet.set_column(0, 0, 32)
        worksheet.set_column(1, 1, 10)
        worksheet.set_column(2, 2, 32)
        worksheet.set_column(3, 3, 72)
        for col in range(4, 9):
            worksheet.set_column(col, col, 30)

        worksheet.freeze_panes(2, 0)

        # заполнение заголовка
        header_fmt = workbook.add_format(formatter.get_header_format(font_size=18))
        base_fmt = workbook.add_format(formatter.get_base_format(font_size=18))
        base_fmt_bold = workbook.add_format(formatter.get_base_format(font_size=18, bold=True))
        base_fmt_bold_red = workbook.add_format(formatter.get_base_format(font_size=18, bold=True, font_color='red'))
        first_row_text = 'Staffing vs Charging report\n{} - {}'.format(date_from.strftime("%d.%m"),
                                                                       date_to.strftime("%d.%m.%Y"))
        worksheet.merge_range('A1:I1', first_row_text, header_fmt)

        row = 1
        col = 0
        for name in ['Specialist', 'Grade', 'Counselor', 'Staffing',
                    'Staffing (total)', 'Project manager', 'Charged on client codes',
                    'Charging - Staffing', 'Comment']:
            worksheet.write(row, col, name, header_fmt)
            col += 1

        row = 2
        for gpn in report.index:

            col = 0

            # вывод специалиста
            worksheet.write(row, col, report.loc[gpn, 'Name'], base_fmt_bold)
            col += 1

            # вывод грейда
            worksheet.write(row, col, report.loc[gpn, 'Short Grade'], base_fmt)
            col += 1

            # вывод канселора
            worksheet.write(row, col, report.loc[gpn, 'Counselor'], base_fmt)
            col += 1

            # вывод стаффинга
            text = staffing_cell_generator.get_cell_text(gpn, date_from)
            format = workbook.add_format(staffing_cell_generator.get_cell_format(gpn, date_from))
            worksheet.write(
                row,
                col,
                text,
                format
            )
            col += 1

            # вывод стаффинга (тотал)
            worksheet.write(row, col, report.loc[gpn, 'Staffing (total)'], base_fmt)
            col += 1

            # вывод пустой колонки с манагерами
            worksheet.write(row, col, report.loc[gpn, 'Project manager'], base_fmt)
            col += 1

            # вывод чарджинга
            worksheet.write(row, col, report.loc[gpn, 'Charged on client codes'], base_fmt)
            col += 1

            # вывод разницы
            diff = float(report.loc[gpn, 'Charging - Staffing'])
            diff = round(diff, 2)

            if (diff < 0):
                fmt = base_fmt_bold_red
                res = str(diff)

            if (diff == 0):
                fmt = base_fmt_bold
                res = str(diff)

            if (diff > 0):
                fmt = base_fmt_bold
                res = "+" + str(diff)

            res = res.replace('.0', '')

            worksheet.write(row, col, res, fmt)
            col += 1

            # вывод комментария
            comment_text = 'Vacation' if 'Vacation' in text else ''
            worksheet.write(row, col, comment_text, base_fmt)
            col += 1

            # вывод одной строки закончен, переходим к следующей
            row += 1

        workbook.close()
        self.result_msg = {'status': 'ok', 'message': f'Отчет сохранен в файл {output_file_name}'}

class CellFormatter:

    def __init__(self, fmt_type=1) -> None:
        with open('formats.json', 'r') as f:
            self.color_ranges = json.load(f)[str(fmt_type)]
        
        self.base_format = {
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'text_wrap': True
        }

        self.colors = {
            'white': '#FFFFFF',
            'yellow': '#FFFF00',
            'green': '#90ee90',
            'red': '#ff5050',
            'bordo': '#b00000',
            'dark_gray': '#565656',
            'header_gray': '#d9d9d9'
        }

    def get_staffing_cell_format(self, total_hours):

        format = self.base_format.copy()

        for color, rng in self.color_ranges.items():
            if rng[0] <= total_hours < rng[1]:
                color_match = color
        
        try:
            format['bg_color'] = self.colors[color_match]
            if color_match in ('bordo', 'dark_gray'):
                format['font_color'] = 'white'
        except UnboundLocalError:
            raise Exception('В файле formats.json есть разрыв периода')

        return format

    def get_header_format(self, font_size=None):

        format = self.base_format.copy()

        if font_size is not None:
            format['font_size'] = font_size

        format['bg_color'] = self.colors['header_gray']
        format['bold'] = True

        return format
    
    def get_base_format(self, font_size=None, bold=False, font_color=None):

        format = self.base_format.copy()

        if font_size is not None:
            format['font_size'] = font_size

        if bold:
            format['bold'] = True

        if font_color is not None:
            format['font_color'] = font_color

        return format

class StaffingReportCellGenerator:

    def __init__(self, staffing_df, fmt_type) -> None:
        self.staffing_df = staffing_df
        self.formatter = CellFormatter(fmt_type)            

    def get_cell_text(self, gpn, week):
        df_filtered = self.staffing_df.loc[(self.staffing_df['GPN'] == gpn) & (self.staffing_df['Период'] == week)]
        job_hours_df = df_filtered[['Job', 'Hours']].groupby('Job', as_index=False).sum()
        hours_list = [hours for _, hours in job_hours_df.values.tolist()]
        
        text = str()
        if job_hours_df.empty or not(any(hours_list)):
            text = '="0"'
        else:
            text = '='
            for job_name, hours in job_hours_df.values.tolist():
                if hours != 0:
                    text += f'"{job_name} ({hours:.0f})"&char(10)&'
                
            text = text[:-10]
        
        return text

    def get_cell_total(self, gpn, week):
        df_filtered = self.staffing_df.loc[(self.staffing_df['GPN'] == gpn) & (self.staffing_df['Период'] == week)]
        staff_hours_df = df_filtered[['GPN', 'Hours']].groupby('GPN', as_index=False).sum()
        total = 0 if staff_hours_df.empty else staff_hours_df['Hours'].values[0]
        return total

    def get_cell_format(self, gpn, week):
        total = self.get_cell_total(gpn, week)
        return self.formatter.get_staffing_cell_format(total)
        
class EmployeeDataLoader:

    def __init__(self, path_to_file) -> None:
        grades_df = pd.read_excel(path_to_file, sheet_name='Grades', index_col=None)
        grades = {v[0]: v[1] for v in grades_df.values}
        grades_order = {key: n for (n, key) in enumerate(grades.keys())}
        
        data_df = pd.read_excel(path_to_file, converters={'GPN': str}, sheet_name='Data')
        data_df.set_index('GPN', inplace=True)
        data_df['Short Grade'] = data_df['Grade'].map(grades)
        data_df['grade_order'] = data_df['Grade'].map(grades_order)
        data_df.sort_values(by=['grade_order', 'Name'], inplace=True)
        data_df[['Counselor']] = data_df[['Counselor']].fillna(value='')

        self.df = data_df

    def get_employee_df(self):
        return self.df

    def get_employee_list(self):
        df = self.df[['Name', 'Short Grade']].reset_index()
        return df.values.tolist()

class StaffingDataLoader:

    def __init__(self, path_to_file, date_from=None, date_to=None):
        self.data_path = path_to_file
        self.load_data()
        self.preprocess_data()
        if date_from:
            self.remove_data_before(date_from)
        if date_to:
            self.remove_data_after(date_to)

    def load_data(self):
        self.raw_df = pd.read_excel(
            self.data_path,
            converters={
                'Период': lambda x: datetime.strptime(x, "%d.%m.%Y").date(),
                'GPN': str,
                'MU': str
            }
        )

    def preprocess_data(self):
        df = self.raw_df.copy()
        df['Job'] = df['Job'].str.strip()
        df['Position'] = df['Position'].str.strip()
        df['Position'] = df['Position'].fillna('')
        df['Staff'] = df['Staff'].str.replace(', ', ' ')
        df = df[df['Staff.Suspended'] == 'Нет']
        df = df[df['MU'] == '00217']
        df = df[['GPN', 'Период', 'Job', 'Hours']]
        df['Период'] = df['Период'] + timedelta(days=2)
        self.df = df

    def get_week_cols(self):
        week_cols = self.df['Период'].unique().tolist()
        week_cols.sort()
        return week_cols

    def remove_data_before(self, date_from):
        self.df = self.df[self.df['Период'] >= date_from]

    def remove_data_after(self, date_to):
        self.df = self.df[self.df['Период'] <= date_to]

    def get_total_df(self, date_from, date_to):
        df = self.df.copy()
        if date_from is not None:
            df = df[df['Период'] >= date_from]
        if date_to is not None:
            df = df[df['Период'] <= date_to]
        df = df.groupby('GPN').sum(numeric_only=True)
        df = df.rename({'Hours': 'Staffing (total)'}, axis=1)
        return df

    def get_df(self):
        return self.df

class ChargingDataLoader:

    def __init__(self, file_path):
        self.file_path = file_path
        self.load_data()
        self.preprocess_data()
        self.filter_data()


    def load_data(self):
        self.raw_df = pd.read_excel(self.file_path,
                                    sheet_name='Details',
                                    skiprows=5,
                                    index_col=None,
                                    converters={'GPN': str})

    def preprocess_data(self):
        df = self.raw_df.copy()
        df.columns = df.columns.str.replace('\n','')
        df['Timesheet Date'] = df['Timesheet Date'].dt.date
        self.df = df

    def filter_data(self):
        self.df = self.df[self.df['Eng. Type'] == 'C']
        self.df = self.df[['GPN', 'Hrs', 'Timesheet Date']]
   
    def get_total_df(self, date_from=None, date_to=None):
        df = self.df.copy()
        if date_from is not None:
            df = df[df['Timesheet Date'] >= date_from]
        if date_to is not None:
            df = df[df['Timesheet Date'] <= date_to]
        df = df.groupby('GPN').sum(numeric_only=True)
        df = df.rename({'Hrs': 'Charged on client codes'}, axis=1)
        return df


if __name__ == '__main__':
    view = View()
    view.main()