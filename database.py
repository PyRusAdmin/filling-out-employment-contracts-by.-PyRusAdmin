# -*- coding: utf-8 -*-
import sqlite3

import openpyxl as op
from loguru import logger

async def import_excel_to_db():
    file = 'list_gup/Списочный_состав.xlsx'
    wb = op.load_workbook(file)
    ws = wb.active

    for row in ws.iter_rows(min_row=5, max_row=1082, min_col=0, max_col=40):
        row_data = [cell.value for cell in row]

        # Подключаемся к базе данных SQLite (или создаем её, если она не существует)
        conn = sqlite3.connect('contracts.db')
        cursor = conn.cursor()

        # Создаем таблицу (если она еще не существует)
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            a0, a1, a2, a3, a4, a5_табельный_номер, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26_номер_договора, a27, a28, a29, a30, a31, a32, a33, a34, a35
        )
        ''')
        # Записываем данные в таблицу employees
        cursor.execute('''
            INSERT INTO employees (a0, a1, a2, a3, a4, a5_табельный_номер, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26_номер_договора, a27, a28, a29, a30, a31, a32, a33, a34, a35)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (row_data[0], row_data[1], row_data[2], row_data[3], row_data[4], row_data[5], row_data[6],
              row_data[7], row_data[8], row_data[9], row_data[10], row_data[11], row_data[12], row_data[13],
              row_data[14], row_data[15], row_data[16], row_data[17], row_data[18], row_data[19], row_data[20],
              row_data[21], row_data[22], row_data[23], row_data[24], row_data[25], row_data[26], row_data[27],
              row_data[28], row_data[29], row_data[30], row_data[31], row_data[32], row_data[33], row_data[34],
              row_data[35]))

        conn.commit()
    logger.info("Данные из Excel импортированы в базу данных.")


async def read_from_db():
    """Считываем данные из базы данных"""

    # Подключаемся к базе данных SQLite
    conn = sqlite3.connect('contracts.db')
    cursor = conn.cursor()
    # Выполняем SQL-запрос для чтения всех данных из таблицы employees
    cursor.execute('SELECT * FROM employees')
    # Получаем все строки результата запроса
    rows = cursor.fetchall()
    # Закрываем соединение с базой данных
    conn.close()
    # Возвращаем данные
    return rows


if __name__ == '__main__':
    import_excel_to_db()