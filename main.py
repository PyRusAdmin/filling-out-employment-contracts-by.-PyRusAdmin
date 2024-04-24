from datetime import datetime
import sqlite3
import openpyxl as op
from docxtpl import DocxTemplate
from gigachat import GigaChat
import configparser
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook

config = configparser.ConfigParser()
config.read('config.ini')
token = config['token']['token']


def gigachat(employee_name):
    # Используйте токен, полученный в личном кабинете из поля Авторизационные данные
    with GigaChat(credentials=token, verify_ssl_certs=False) as giga:
        response = giga.chat(f"Определи по Имени Фамилии и Отчеству это мужчина или женщина? "
                             f"Ответ дай короткий или мужчина или женщина: {employee_name}")
        return response.choices[0].message.content


def opening_a_file():
    """Открытие файла Excel"""
    root = Tk()
    root.withdraw()
    filename = askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    return filename


def opening_the_database():
    """Открытие базы данных"""
    conn = sqlite3.connect('data.db')  # Создаем соединение с базой данных
    cursor = conn.cursor()
    return conn, cursor


def po_parsing_jul_2023():
    """Изменение от 19.01.2024 Парсинг май 2023"""

    conn, cursor = opening_the_database()
    filename = opening_a_file()  # Открываем выбор файла Excel для чтения данных
    workbook = load_workbook(filename=filename)  # Загружаем выбранный файл Excel
    sheet = workbook.active
    # Создаем таблицу в базе данных, если она еще не существует
    cursor.execute(
        '''CREATE TABLE IF NOT EXISTS parsing (service_number, full_name, phone, address, series_number, issue_date, issued_by, code)''')
    # Считываем данные из колонок A и H и вставляем их в базу данных
    for row in sheet.iter_rows(min_row=5, max_row=1076, values_only=True):
        service_number = str(row[0])  # Преобразуем значение в строку
        full_name = str(row[1])
        phone = str(row[2])
        address = str(row[3])
        series_number = str(row[4])
        issue_date = str(row[5])
        issued_by = str(row[6])
        code = str(row[7])
        print(service_number, full_name, phone, address, series_number, issue_date, issued_by, code)
        # Проверяем, существует ли запись с таким табельным номером в базе данных
        cursor.execute('SELECT * FROM parsing WHERE service_number = ?', (service_number,))
        existing_row = cursor.fetchone()
        # Если запись с таким табельным номером не существует, вставляем данные в базу данных
        if existing_row is None:
            cursor.execute('INSERT INTO parsing VALUES (?, ?, ?,?, ?,?,?,?)',
                           (service_number, full_name, phone, address, series_number, issue_date, issued_by, code,))
        # Сохраняем изменения в базе данных и закрываем соединение
    conn.commit()
    conn.close()


def comparing_property():
    """Сравниваем данные с базы данных с файлом"""
    conn, cursor = opening_the_database()
    # Загружаем файл Excel для записи результатов
    result_workbook = load_workbook(filename='list_gup/Списочный_состав.xlsx')
    result_sheet = result_workbook.active
    cursor.execute(
        'SELECT service_number, full_name, phone, address, series_number, issue_date, issued_by, code FROM parsing')  # Получаем все данные из базы данных
    db_data = cursor.fetchall()  # Получаем все записи из базы данных
    # Сравниваем значения в колонке D с базой данных и записываем результаты в колонки G, H и I
    for row in result_sheet.iter_rows(min_row=6, max_row=1077):
        value_D = str(row[5].value)  # Значение в колонке D
        print(value_D)
        db_number_list = [db_row for db_row in db_data if db_row[0] == value_D]
        print(db_number_list)
        if db_number_list:
            full_name = db_number_list[0][1]
            print("Found full name:", full_name)
            row[17].value = full_name
            phone = db_number_list[0][2]
            row[18].value = phone
            print(phone)
            address = db_number_list[0][3]
            row[19].value = address
            series_number = db_number_list[0][4]
            row[20].value = series_number  # Год из базы данных в колонку 20
            issue_date = db_number_list[0][5]
            row[21].value = issue_date
            issued_by = db_number_list[0][6]
            row[22].value = issued_by
            code = db_number_list[0][7]
            row[23].value = code  # Год из базы данных в колонку 20

    # Сохраняем изменения в файле Excel для записи результатов
    result_workbook.save(filename='list_gup/Списочный_состав.xlsx')
    result_workbook.close()


def open_list_gup_docx():
    """Открываем документ с шаблоном"""
    wb = op.load_workbook('list_gup/Списочный_состав_1.xlsx')  # открываем файл
    sheet = wb.active  # открываем активную таблицу
    current_row = 15  # Начальная строка
    for row in sheet.iter_rows(min_row=6, max_row=1077, values_only=True):
        number = str(row[6])  # Считываем значение в колонке
        sheet.cell(row=current_row, column=27, value=gigachat(number))  # Записываем значение в ячейку
        current_row += 1  # Переходим к следующей строке
    wb.save("list_gup/Списочный_состав_1.xlsx")


def open_list_gup():
    # file = 'list_gup/Список.xlsx'
    file = 'list_gup/Списочный_состав.xlsx'
    wb = op.load_workbook(file)  # открываем файл
    ws = wb.active  # открываем активную таблицу
    list_gup = []  # создаем список
    for row in ws.iter_rows(min_row=6, max_row=1077, min_col=0, max_col=23):  # перебираем строки
        row_data = [cell.value for cell in row]  # создаем список
        list_gup.append(row_data)  # добавляем в список
    return list_gup  # возвращаем список


def filling_data_hourly_rate(row, formatted_date, ending, file_dog):
    doc = DocxTemplate(file_dog)
    context = {
        'name_surname': f" {row[6]} ",  # Ф.И.О. (Иванов Иван Иванович)
        'name_surname_completely': f" {row[7]} ",  # Ф.И.О. (Иванов И. И.)
        'date_admission': f" {formatted_date} ",  # Дата поступления
        'ending': f"{ending}",  # Окончание ый или ая
        'post': f" {row[3]} ",  # Должность
        'district': f" {row[1]} ",  # Участок
        'salary': f" {row[11]} ",  # Часовая тарифная ставка или оклад
        'series_number': f'{row[17]}',  # Номер паспорта
        'phone': f'{row[15]}',  # Телефон
        'address': f'{row[16]}',  # Адрес
        'issue_date': f'{row[18]}',  # Дата выдачи
        'issued_by': f'{row[19]}',  # Кем выдан
        'code': f'{row[20]}',  # Код подразделения
        'official_salary': 'часовая тарифная ставка',
        'official_salary_termination': 'часовой тарифной ставки',
        'month_or_hour': 'в час',
        'district_pro': f" {row[22]} ",  # Участок
    }
    doc.render(context)
    doc.save(f"готовые договора/{row[0]}_{row[5]}_{row[6]}.docx")


def record_data_salary(row, formatted_date, ending, file_dog):
    doc = DocxTemplate(file_dog)
    context = {
        'name_surname': f" {row[6]} ",  # Ф.И.О. (Иванов Иван Иванович)
        'name_surname_completely': f" {row[7]} ",  # Ф.И.О. (Иванов И. И.)
        'date_admission': f" {formatted_date} ",  # Дата поступления
        'ending': f"{ending}",  # Окончание ый или ая
        'post': f" {row[3]} ",  # Должность
        'district': f" {row[1]} ",  # Участок
        'salary': f" {row[11]} ",  # Часовая тарифная ставка или оклад
        'series_number': f'{row[17]}',  # Номер паспорта
        'phone': f'{row[15]}',  # Телефон
        'address': f'{row[16]}',  # Адрес
        'issue_date': f'{row[18]}',  # Дата выдачи
        'issued_by': f'{row[19]}',  # Кем выдан
        'code': f'{row[20]}',  # Код подразделения
        'official_salary': 'должностной оклад',
        'official_salary_termination': 'должностного оклада',
        'month_or_hour': 'в месяц',
        'district_pro': f" {row[22]} ",  # Участок
    }
    doc.render(context)
    doc.save(f"готовые договора/{row[0]}_{row[5]}_{row[6]}.docx")


def creation_contracts(row, formatted_date, ending):

    if row[11] > 1000:
        if row[21] == 7:
            file_dog = "template/Шаблон_трудовой_договор_7_часов.docx"
            record_data_salary(row, formatted_date, ending, file_dog)
        elif row[21] == 8:
            if row[2] == 'Рук.пр.гр.подз':
                file_dog = "template/Шаблон_трудовой_договор_8_часов_ИТР.docx"
                record_data_salary(row, formatted_date, ending, file_dog)
            elif row[2] == 'Спец.пром.подз':
                file_dog = "template/Шаблон_трудовой_договор_8_часов_ИТР.docx"
                record_data_salary(row, formatted_date, ending, file_dog)
            else:
                file_dog = "template/Шаблон_трудовой_договор.docx"
                record_data_salary(row, formatted_date, ending, file_dog)
        else:
            file_dog = "template/Шаблон_трудовой_договор.docx"
            record_data_salary(row, formatted_date, ending, file_dog)
    elif row[11] < 1000:
        if row[21] == 6:
            file_dog = "template/Шаблон_трудовой_договор_6_часов.docx"
            filling_data_hourly_rate(row, formatted_date, ending, file_dog)
        else:
            file_dog = "template/Шаблон_трудовой_договор.docx"
            filling_data_hourly_rate(row, formatted_date, ending, file_dog)


def format_date(date):
    months = {
        1: "января", 2: "февраля", 3: "марта", 4: "апреля",
        5: "мая", 6: "июня", 7: "июля", 8: "августа",
        9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"
    }
    date = datetime.strptime(date, "%d.%m.%Y")
    return '" {:02d} " {} {} г.'.format(date.day, months[date.month], date.year)


if __name__ == '__main__':
    parsed_data = open_list_gup()
    for row in parsed_data:
        print(row)
        if row[14] == "Мужчина":
            ending = "ый"
            formatted_date = format_date(row[8])
            creation_contracts(row, formatted_date, ending)
        elif row[14] == "Женщина":
            ending = "ая"
            formatted_date = format_date(row[8])
            creation_contracts(row, formatted_date, ending)
