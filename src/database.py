# -*- coding: utf-8 -*-
import openpyxl as op
from loguru import logger
from peewee import *

# Настройка базы данных через Peewee
db = SqliteDatabase("contracts.db")


class Employee(Model):
    a0 = CharField(null=True)
    a1 = CharField(null=True)
    a2 = CharField(null=True)
    a3 = CharField(null=True)
    a4_табельный_номер = CharField(null=True)
    a5 = CharField(null=True)
    a6 = CharField(null=True)
    a7 = CharField(null=True)  # Дата поступление на предприятие
    a8 = CharField(null=True)
    a9 = CharField(null=True)
    a10 = CharField(null=True)
    a11 = CharField(null=True)  # Гендер сотрудника (мужской/женский)
    a12 = CharField(null=True)
    a13 = CharField(null=True)
    a14 = CharField(null=True)
    a15 = CharField(null=True)
    a16 = CharField(null=True)
    a17 = CharField(null=True)
    a18 = CharField(null=True)
    a19 = CharField(null=True)
    a20 = CharField(null=True)
    a21 = CharField(null=True)
    a22 = CharField(null=True)
    a23 = CharField(null=True)
    a24 = CharField(null=True)
    a25_номер_договора = CharField(null=True)
    a26 = CharField(null=True)
    a27 = CharField(null=True)
    a28 = CharField(null=True)
    a29 = CharField(null=True)
    a30 = CharField(null=True)
    a31 = CharField(null=True)  # напечатанный
    a32 = CharField(null=True)
    a33 = CharField(null=True)
    a34 = CharField(null=True)

    class Meta:
        database = db  # Указываем, что модель будет использовать нашу базу данных


# Функция для импорта данных из Excel в базу данных
async def import_excel_to_db(min_row, max_row):
    file = "../data/list_gup/Списочный_состав.xlsx"
    wb = op.load_workbook(file)
    ws = wb.active

    # Подключаемся к базе данных и создаем таблицу, если она не существует
    db.connect()
    db.create_tables([Employee], safe=True)

    # Импортируем данные
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=0, max_col=40):
        row_data = [cell.value for cell in row]

        # Создаем запись в базе данных
        Employee.create(
            a0=row_data[0], a1=row_data[1], a2=row_data[2], a3=row_data[3],
            a4_табельный_номер=row_data[4], a5=row_data[5], a6=row_data[6],
            a7=row_data[7], a8=row_data[8], a9=row_data[9], a10=row_data[10],
            a11=row_data[11], a12=row_data[12], a13=row_data[13], a14=row_data[14],
            a15=row_data[15], a16=row_data[16], a17=row_data[17], a18=row_data[18],
            a19=row_data[19], a20=row_data[20], a21=row_data[21], a22=row_data[22],
            a23=row_data[23], a24=row_data[24], a25_номер_договора=row_data[25],
            a26=row_data[26], a27=row_data[27], a28=row_data[28], a29=row_data[29],
            a30=row_data[30], a31=row_data[31], a32=row_data[32], a33=row_data[33],
            a34=row_data[34],
        )

    db.close()  # Закрываем подключение к базе данных
    logger.info("Данные из Excel импортированы в базу данных.")


async def read_from_db():
    """Функция для чтения данных из базы данных. Считываем данные из базы данных"""
    db.connect()
    rows = Employee.select()  # Получаем все записи из таблицы employees
    db.close()  # Закрываем подключение к базе данных
    return rows


# Функция для очистки базы данных
async def clear_database():
    """Удаляет все записи из таблицы Employee."""
    try:
        db.connect()
        deleted_count = Employee.delete().execute()
        db.close()
        logger.info(f"База данных очищена. Удалено записей: {deleted_count}")
    except Exception as e:
        logger.exception("Ошибка при очистке базы данных: ", e)


if __name__ == "__main__":
    import_excel_to_db()
    clear_database()  # Очистка базы данных
