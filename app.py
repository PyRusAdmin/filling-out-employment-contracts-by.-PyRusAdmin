import time
from datetime import datetime
from threading import Thread

from flask import Flask, render_template, request, redirect, url_for
from loguru import logger

from database import import_excel_to_db, read_from_db
from filling_data import creation_contracts, format_date
from parsing_comparison_file import parsing_document_1, compare_and_rewrite_professions

app = Flask(__name__)


progress_messages = []  # список сообщений, которые будут отображаться в progress

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/loading')
def loading():
    """Сообщение, что база данных формируется"""
    return render_template('loading.html')

def run_import():
    # Импорт данных из Excel в базе данных в фоне
    import_excel_to_db()


@app.route('/progress')
def progress():
    def generate():
        while True:
            if progress_messages:
                message = progress_messages.pop(0)
                yield f"data: {message}\n\n"
            time.sleep(1)


@app.route('/action', methods=['POST'])
def action():
    user_input = int(request.form['user_input'])

    if user_input == 1:
        parsing_document_1(min_row=6, max_row=1084, column=5, column_1=8)
    elif user_input == 2:
        start = datetime.now()  # фиксируем и выводим время старта работы кода
        logger.info('Время старта: ' + str(start))
        data = read_from_db() # Считываем данные из базы данных
        # Выводим считанные данные на экран
        for row in data:
            logger.info(row)
            if row[14] == "Мужчина":
                ending = "ый"
                logger.info(row[8])  # выводим данные в консоль
                creation_contracts(row, format_date(row[8]), ending)
            elif row[14] == "Женщина":
                ending = "ая"
                creation_contracts(row, format_date(row[8]), ending)
        finish = datetime.now()  # фиксируем и выводим время окончания работы кода
        logger.info('Время окончания: ' + str(finish))
        logger.info('Время работы: ' + str(finish - start))  # вычитаем время старта из времени окончания

    elif user_input == 3:
        compare_and_rewrite_professions()


    elif user_input == 4:

        # Запускаем фоновый поток для импорта данных и сразу отображаем страницу загрузки
        thread = Thread(target=run_import)
        thread.start()
        return redirect(url_for('loading'))

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)