from peewee import *

db = SqliteDatabase("contracts.db")  # Создайте модель для таблицы в базе данных


class Employee(Model):
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
