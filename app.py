import os
from flask import Flask, request, render_template, jsonify
import pygsheets
import pandas as pd
from openpyxl import load_workbook
from win32com.client import Dispatch as client
import threading
import pickle
import pythoncom

app = Flask(__name__)

# Функция для конвертации xls файла в xlsx
def convert_xls_to_xlsx(xls_file):
    try:
        pythoncom.CoInitialize()  # Инициализация COM-библиотек
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
    finally:
        pythoncom.CoUninitialize()  # Освобождение COM-библиотек

# Функция для загрузки данных в Google Sheets
def upload_to_google_sheets(gc, selected_sheet, selected_tab, excel_file):
    try:
        if excel_file.endswith('.xls'):
            converted_file = convert_xls_to_xlsx(excel_file)
            if converted_file:
                excel_file = converted_file
            else:
                return {"status": "error", "message": "Ошибка при конвертации файла. Загрузка отменена."}
        
        sh = gc.open(selected_sheet)
        worksheet = sh.worksheet_by_title(selected_tab)

        if os.path.exists(excel_file):
            wb = load_workbook(excel_file)
            ws = wb.active
            df = pd.DataFrame(ws.values)
            max_rows, max_cols = min(df.shape[0], 1086), min(df.shape[1], 56)
            df_selected = df.iloc[:max_rows, :max_cols]
            worksheet.clear()
            worksheet.update_values(crange='A1', values=df_selected.values.tolist())
            
            if excel_file.endswith('.xlsx'):
                os.remove(excel_file)
            return {"status": "success", "message": "Данные успешно загружены в Google Sheets."}
        else:
            return {"status": "error", "message": "Файл Excel не найден."}
    except Exception as e:
        print(f"Ошибка при загрузке данных: {e}")
        return {"status": "error", "message": f"Ошибка при загрузке данных: {e}"}

# Функция для запуска загрузки данных в отдельном потоке
def start_upload_thread(gc, selected_sheet, selected_tab, excel_file):
    result = {"status": "error", "message": "Файл Excel не выбран."}
    if excel_file:
        def target():
            nonlocal result
            result = upload_to_google_sheets(gc, selected_sheet, selected_tab, excel_file)
        
        thread = threading.Thread(target=target)
        thread.daemon = True
        thread.start()
        thread.join()  # Ждем завершения потока
    return result

# Проверяем наличие сохраненного объекта gc
try:
    with open('gc.pickle', 'rb') as f:
        gc = pickle.load(f)
except FileNotFoundError:
    credentials_file = 'credentials.json'
    gc = pygsheets.authorize(service_account_file=credentials_file)
    with open('gc.pickle', 'wb') as f:
        pickle.dump(gc, f)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_sheets', methods=['GET'])
def get_sheets():
    sheets = gc.spreadsheet_titles()
    return jsonify(sheets)

@app.route('/get_tabs', methods=['POST'])
def get_tabs():
    sheet_name = request.json['sheet_name']
    sh = gc.open(sheet_name)
    tabs = [ws.title for ws in sh.worksheets()]
    return jsonify(tabs)

@app.route('/upload', methods=['POST'])
def upload():
    selected_sheet = request.form['selected_sheet']
    selected_tab = request.form['selected_tab']
    excel_file = request.files['excel_file']
    
    file_path = os.path.join(os.getcwd(), excel_file.filename)
    excel_file.save(file_path)
    
    result = start_upload_thread(gc, selected_sheet, selected_tab, file_path)
    return jsonify(result)

if __name__ == '__main__':
    app.run(debug=True)
