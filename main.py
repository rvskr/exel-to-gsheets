import os
import pandas as pd
import pygsheets
from openpyxl import load_workbook
from tkinter import filedialog, Tk, Label, Button, OptionMenu, StringVar, ttk
from win32com.client import Dispatch as client
from tkinter import messagebox
import threading

# Функция для выбора Google Sheets таблицы
def select_google_sheet(gc, var_sheet, var_tab, tab_menu):
    available_sheets = gc.spreadsheet_titles()
    var_sheet.set(available_sheets[0])

    def on_option_change(*args):
        selected_sheet = var_sheet.get()
        select_google_sheet_tab(gc, selected_sheet, var_tab, tab_menu)

    option_menu = OptionMenu(root, var_sheet, *available_sheets)
    option_menu.grid(row=2, column=1, pady=5)
    var_sheet.trace('w', on_option_change)

# Функция для выбора листа в Google Sheets таблице
def select_google_sheet_tab(gc, selected_sheet, var_tab, tab_menu):
    sh = gc.open(selected_sheet)
    available_tabs = [sheet.title for sheet in sh.worksheets()]
    var_tab.set(available_tabs[0])
    tab_menu['menu'].delete(0, 'end')
    for tab in available_tabs:
        tab_menu['menu'].add_command(label=tab, command=lambda value=tab: var_tab.set(value))

# Функция для выбора файла Excel
def select_excel_file(var, file_label):
    filename = filedialog.askopenfilename()
    var.set(filename)
    file_label.config(text=f"Выбранный файл: {filename}")

# Функция для конвертации xls файла в xlsx
def convert_xls_to_xlsx(xls_file):
    try:
        file_path, file_name = os.path.split(xls_file)
        file_name_without_extension = os.path.splitext(file_name)[0]
        xlsx_file = os.path.join(file_path, f"{file_name_without_extension}.xlsx")
        xlsx_file = os.path.normpath(xlsx_file)
        
        if os.path.exists(xlsx_file):
            os.remove(xlsx_file)

        excel = client("Excel.Application")
        wb = excel.Workbooks.Open(xls_file)
        wb.SaveAs(xlsx_file, FileFormat=51)
        wb.Close()
        excel.Quit()
        
        return xlsx_file
    except Exception as e:
        print(f"Ошибка при конвертации файла: {e}")
        return None

# Функция для загрузки данных в Google Sheets
def upload_to_google_sheets(gc, selected_sheet, selected_tab, excel_file):
    try:
        if excel_file.get().endswith('.xls'):
            converted_file = convert_xls_to_xlsx(excel_file.get())
            if converted_file:
                excel_file.set(converted_file)
            else:
                messagebox.showerror("Ошибка", "Ошибка при конвертации файла. Загрузка отменена.")
                return False
        
        sh = gc.open(selected_sheet.get())
        worksheet = sh.worksheet_by_title(selected_tab.get())

        if os.path.exists(excel_file.get()):
            wb = load_workbook(excel_file.get())
            ws = wb.active
            df = pd.DataFrame(ws.values)
            max_rows, max_cols = min(df.shape[0], 1086), min(df.shape[1], 56)
            df_selected = df.iloc[:max_rows, :max_cols]
            worksheet.clear()
            worksheet.update_values(crange='A1', values=df_selected.values.tolist())
            messagebox.showinfo("Успех", "Данные успешно загружены в Google Sheets.")
            
            if excel_file.get().endswith('.xlsx'):
                os.remove(excel_file.get())
                excel_file.set("")
            return True
        else:
            messagebox.showerror("Ошибка", "Файл Excel не найден.")
            return False
    except Exception as e:
        print(f"Ошибка при загрузке данных: {e}")
        messagebox.showerror("Ошибка", f"Ошибка при загрузке данных: {e}")
        return False

# Функция для запуска загрузки данных в отдельном потоке
def start_upload_thread(gc, selected_sheet, selected_tab, excel_file):
    if excel_file.get():
        thread = threading.Thread(target=lambda: upload_to_google_sheets(gc, selected_sheet, selected_tab, excel_file))
        thread.daemon = True
        thread.start()
    else:
        messagebox.showerror("Ошибка", "Файл Excel не выбран.")

# Аутентификация и создание клиента для доступа к API Google Sheets
credentials_file = 'credentials.json'
gc = pygsheets.authorize(service_account_file=credentials_file)

# Создаем окно
root = Tk()
root.title("Загрузка данных в Google Sheets")
root.geometry("500x300")

# Переменные для хранения выбранной таблицы, листа и файла Excel
selected_sheet = StringVar()
selected_tab = StringVar()
excel_file = StringVar()

# Определяем стиль для кнопок
style = ttk.Style()
style.configure('TButton', font=('Helvetica', 12))

# Создаем и размещаем виджеты
header_label = ttk.Label(root, text="Загрузка данных в Google Sheets", font=('Helvetica', 18, 'bold'))
header_label.grid(row=0, column=0, columnspan=2, pady=10)

file_label = ttk.Label(root, text="Выбранный файл: Не выбран", wraplength=300)
file_label.grid(row=1, column=0, padx=10, pady=5)

file_button = ttk.Button(root, text="Выбрать файл", command=lambda: select_excel_file(excel_file, file_label))
file_button.grid(row=1, column=1, padx=10, pady=5)

sheet_label = ttk.Label(root, text="Выберите Google Sheets таблицу:")
sheet_label.grid(row=2, column=0, padx=10, pady=5)

tab_label = ttk.Label(root, text="Выберите лист в таблице:")
tab_label.grid(row=3, column=0, padx=10, pady=5)

tab_menu = OptionMenu(root, selected_tab, "")
tab_menu.grid(row=3, column=1, padx=10, pady=5)

select_google_sheet(gc, selected_sheet, selected_tab, tab_menu)

upload_button = ttk.Button(root, text="Загрузить данные", command=lambda: start_upload_thread(gc, selected_sheet, selected_tab, excel_file))
upload_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()
