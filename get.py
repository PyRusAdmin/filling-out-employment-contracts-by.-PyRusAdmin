from peewee import *

# Создайте модель для таблицы в базе данных
db = SqliteDatabase('contracts.db')


class Employee(Model):
    # id = AutoField()  # Добавляем поле id, которое будет автоматически увеличиваться
    a0 = CharField()
    a1 = CharField()
    a2 = CharField()
    a3 = CharField()
    a4_табельный_номер = CharField()
    a5 = CharField()
    a6 = CharField()
    a7 = CharField()
    a8 = CharField()
    a9 = CharField()
    a10 = CharField()
    a11 = CharField()
    a12 = CharField()
    a13 = CharField()
    a14 = CharField()
    a15 = CharField()
    a16 = CharField()
    a17 = CharField()
    a18 = CharField()
    a19 = CharField()
    a20 = CharField()
    a21 = CharField()
    a22 = CharField()
    a23 = CharField()
    a24 = CharField()
    a25_номер_договора = CharField()
    a26 = CharField()
    a27 = CharField()
    a28 = CharField()
    a29 = CharField()
    a30 = CharField()
    a31 = CharField()
    a32 = CharField()
    a33 = CharField()
    a34 = CharField()

    class Meta:
        database = db
        # primary_key = False  # Отключаем авто-добавление поля id


def search_employee_by_tab_number(tab_number):
    """Ищем данные сотрудника по табельному номеру"""
    try:
        # Поиск строки в базе данных по табельному номеру
        customer = Employee.get(Employee.a4_табельный_номер == tab_number)
        # Возвращаем строку со всеми данными
        return customer
    except Employee.DoesNotExist:
        # Возвращает None, если запись не найдена
        return None


def main():
    tab_number = 22148
    row = search_employee_by_tab_number(tab_number)
    if row:
        # Выводим все поля модели
        print(row.a0, row.a1, row.a2, row.a3, row.a4_табельный_номер, row.a5, row.a6, row.a7, row.a8, row.a9, row.a10,
              row.a11, row.a12, row.a13, row.a14, row.a15, row.a16, row.a17, row.a18, row.a19, row.a20, row.a21, row.a22,
              row.a23, row.a24, row.a25_номер_договора, row.a26, row.a27, row.a28, row.a29, row.a30, row.a31, row.a32,
              row.a33, row.a34)
    else:
        print("Данные не найдены")


if __name__ == '__main__':
    main()

