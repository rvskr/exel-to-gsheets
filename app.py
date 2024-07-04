import os
from flask import Flask, request, render_template, jsonify
import pygsheets
import pandas as pd
import xlrd
import threading
import pickle

app = Flask(__name__)
# Определяем порт для Flask
port = int(os.environ.get('PORT', 5000))

# Функция для загрузки данных из xls файла в DataFrame
def load_data_from_xls(file_path):
    try:
        wb = xlrd.open_workbook(file_path)
        sheet = wb.sheet_by_index(0)
        data = []
        for rownum in range(sheet.nrows):
            data.append(sheet.row_values(rownum))
        return pd.DataFrame(data)
    except Exception as e:
        print(f"Ошибка при загрузке данных из xls файла: {e}")
        return None

# Функция для загрузки данных в Google Sheets
def upload_to_google_sheets(gc, selected_sheet, selected_tab, excel_file):
    try:
        sh = gc.open(selected_sheet)
        worksheet = sh.worksheet_by_title(selected_tab)

        df = load_data_from_xls(excel_file)
        if df is not None:
            max_rows, max_cols = min(df.shape[0], 1086), min(df.shape[1], 56)
            df_selected = df.iloc[:max_rows, :max_cols]
            worksheet.clear()
            worksheet.update_values(crange='A1', values=df_selected.values.tolist())
            return {"status": "success", "message": "Данные успешно загружены в Google Sheets."}
        else:
            return {"status": "error", "message": "Ошибка при загрузке данных из xls файла."}
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
    # Запускаем Flask на указанном порту
    app.run(host='0.0.0.0', port=port)