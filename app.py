from flask import Flask, render_template, request, redirect, url_for
from datetime import datetime

from database import import_excel_to_db
from main import open_list_gup, creation_contracts, format_date
from parsing_comparison_file import parsing_document_1, compare_and_rewrite_professions
from loguru import logger


app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/action', methods=['POST'])
def action():
    user_input = int(request.form['user_input'])

    if user_input == 1:
        parsing_document_1(min_row=6, max_row=1084, column=5, column_1=8)
    elif user_input == 2:

        start = datetime.now()  # фиксируем и выводим время старта работы кода
        logger.info('Время старта: ' + str(start))

        parsed_data = open_list_gup()
        for row in parsed_data:
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

        import_excel_to_db() # импортируем данные из excel в базу данных

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)