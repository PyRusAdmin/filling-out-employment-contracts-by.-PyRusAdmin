# -*- coding: utf-8 -*-
from datetime import datetime
import sqlite3
import openpyxl as op
from docxtpl import DocxTemplate
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook
from loguru import logger

logger.add("log/log.log", rotation="1 MB", compression="zip")  # Логирование программы
table_name = "parsing"  # Имя таблицы в базе данных
file_database = "data.db"  # Имя файла базы данных


def opening_a_file():
    """Открытие файла Excel выбором файла"""
    root = Tk()
    root.withdraw()
    filename = askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    return filename


def opening_the_database():
    """Открытие базы данных"""
    conn = sqlite3.connect('data.db')  # Создаем соединение с базой данных
    cursor = conn.cursor()
    return conn, cursor


def parsing_document_1(min_row, max_row, column, column_1) -> None:
    """
    Осуществляет парсинг данных из файла Excel и вставляет их в базу данных SQLite.

    Аргументы:
    :param min_row: Строка, с которой начинается считывание данных.
    :param max_row: Строка, с которой заканчивается считывание данных.
    :param column: Столбец, с которого начинается считывание данных.
    :param column_1: Столбец, с которого начинается считывание данных.
    """
    filename = opening_a_file()  # Открываем выбор файла Excel для чтения данных
    workbook = load_workbook(filename=filename)  # Загружаем выбранный файл Excel
    sheet = workbook.active

    os.remove(file_database)  # Удаляем файл базы данных

    conn = sqlite3.connect(file_database)  # Создаем соединение с базой данных
    cursor = conn.cursor()
    # Создаем таблицу в базе данных, если она еще не существует
    cursor.execute(f"CREATE TABLE IF NOT EXISTS {table_name} (table_column_1, table_column_2)")
    # Считываем данные из колонки A и вставляем их в базу данных
    for row in sheet.iter_rows(min_row=int(min_row), max_row=int(max_row), values_only=True):
        table_column_1 = str(row[int(column)])  # Преобразуем значение в строку
        table_column_2 = str(row[int(column_1)])  # Преобразуем значение в строку
        # Проверяем, существует ли запись с таким табельным номером в базе данных
        cursor.execute(f"SELECT * FROM {table_name} WHERE table_column_1 = ? AND table_column_2 = ?",
                       (table_column_1, table_column_2))
        existing_row = cursor.fetchone()
        # Если запись с таким табельным номером не существует, вставляем данные в базу данных
        if existing_row is None:
            cursor.execute(f"INSERT INTO {table_name} VALUES (?, ?)", (table_column_1, table_column_2))
    # Удаляем повторы по табельному номеру
    cursor.execute(
        f"DELETE FROM {table_name} WHERE rowid NOT IN (SELECT min(rowid) FROM {table_name} GROUP BY table_column_1, table_column_2)")
    # Сохраняем изменения в базе данных и закрываем соединение
    conn.commit()
    conn.close()


def open_list_gup():
    file = 'list_gup/Списочный_состав.xlsx'
    wb = op.load_workbook(file)  # открываем файл
    ws = wb.active  # открываем активную таблицу
    list_gup = []  # создаем список
    for row in ws.iter_rows(min_row=6, max_row=1077, min_col=0, max_col=36):  # перебираем строки
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
        if row[21] == 7:  # 7 часов
            file_dog = "template/Шаблон_трудовой_договор_7_часов.docx"
            record_data_salary(row, formatted_date, ending, file_dog)
        elif row[21] == 8:  # 8 часов
            if row[2] == 'Рук.пр.гр.подз':
                file_dog = "template/Шаблон_трудовой_договор_8_часов_ИТР.docx"
                record_data_salary(row, formatted_date, ending, file_dog)
            elif row[2] == 'Спец.пром.подз':
                file_dog = "template/Шаблон_трудовой_договор_8_часов_ИТР.docx"
                record_data_salary(row, formatted_date, ending, file_dog)
            else:
                file_dog = "template/Шаблон_трудовой_договор.docx"
                record_data_salary(row, formatted_date, ending, file_dog)
        elif row[21] == 12:  # 12 часов
            print(12)
            file_dog = "template/Шаблон_трудовой_договор_12_часов.docx"
            record_data_salary(row, formatted_date, ending, file_dog)
        else:
            file_dog = "template/Шаблон_трудовой_договор.docx"
            record_data_salary(row, formatted_date, ending, file_dog)
    elif row[11] < 1000:
        if row[21] == 6:  # 6 часов
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
    # TODO: Вынести в отдельную функцию
    parsed_data = open_list_gup()
    for row in parsed_data:
        print(row[32])
        if row[14] == "Мужчина":
            ending = "ый"
            formatted_date = format_date(row[8])
            creation_contracts(row, formatted_date, ending)
        elif row[14] == "Женщина":
            ending = "ая"
            formatted_date = format_date(row[8])
            creation_contracts(row, formatted_date, ending)
