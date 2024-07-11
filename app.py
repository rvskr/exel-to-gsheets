import os
from flask import Flask, request, render_template, jsonify, redirect, url_for, session
import pygsheets
import pandas as pd
import xlrd
import threading
import pickle
from flask_basicauth import BasicAuth
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.config['BASIC_AUTH_USERNAME'] = os.getenv('BASIC_AUTH_USERNAME')
app.config['BASIC_AUTH_PASSWORD'] = os.getenv('BASIC_AUTH_PASSWORD')
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'mysecret')

basic_auth = BasicAuth(app)

port = int(os.environ.get('PORT', 5000))

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

def upload_to_google_sheets(gc, selected_sheet, selected_tab, excel_file):
    try:
        sh = gc.open(selected_sheet)
        worksheet = sh.worksheet_by_title(selected_tab)

        df = load_data_from_xls(excel_file)
        if df is not None:
            max_rows, max_cols = min(df.shape[0], 1086), min(df.shape[1], 56)
            df_selected = df.iloc[:max_rows, :max_cols]
            
            # Удаление пустых строк
            df_selected = df_selected.dropna(how='all')
            
            # Преобразование DataFrame в список списков для загрузки в Google Sheets
            values_to_update = df_selected.values.tolist()
            
            # Очистка листа перед обновлением
            worksheet.clear()
            
            # Загрузка данных в лист
            worksheet.update_values(crange='A1', values=values_to_update)
            
            return {"status": "success", "message": "Данные успешно загружены в Google Sheets."}
        else:
            return {"status": "error", "message": "Ошибка при загрузке данных из xls файла."}
    except Exception as e:
        print(f"Ошибка при загрузке данных: {e}")
        return {"status": "error", "message": f"Ошибка при загрузке данных: {e}"}

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

try:
    with open('gc.pickle', 'rb') as f:
        gc = pickle.load(f)
except FileNotFoundError:
    credentials_content = os.getenv('GOOGLE_SHEETS_CREDENTIALS')
    if credentials_content:
        gc = pygsheets.authorize(service_account_json=credentials_content)
        with open('gc.pickle', 'wb') as f:
            pickle.dump(gc, f)
    else:
        raise ValueError("GOOGLE_SHEETS_CREDENTIALS переменная не найдена в .env файле.")

@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    return render_template('index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        if (request.form['username'] == app.config['BASIC_AUTH_USERNAME'] and
                request.form['password'] == app.config['BASIC_AUTH_PASSWORD']):
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            return 'Неправильный логин или пароль', 401
    return render_template('login.html')

@app.route('/logout', methods=['POST'])
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route('/get_sheets', methods=['GET'])
def get_sheets():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    sheets = gc.spreadsheet_titles()
    return jsonify(sheets)

@app.route('/get_tabs', methods=['POST'])
def get_tabs():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    sheet_name = request.json['sheet_name']
    sh = gc.open(sheet_name)
    tabs = [ws.title for ws in sh.worksheets()]
    return jsonify(tabs)

@app.route('/upload', methods=['POST'])
def upload():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    selected_sheet = request.form['selected_sheet']
    selected_tab = request.form['selected_tab']
    excel_file = request.files['excel_file']
    
    file_path = os.path.join(os.getcwd(), excel_file.filename)
    excel_file.save(file_path)
    
    result = start_upload_thread(gc, selected_sheet, selected_tab, file_path)
    return jsonify(result)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=port)
