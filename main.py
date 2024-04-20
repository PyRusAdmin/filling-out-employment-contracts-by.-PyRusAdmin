from docxtpl import DocxTemplate
import openpyxl as op
from datetime import datetime


def open_list_gup():
    wb = op.load_workbook('list_gup/Списочный_состав.xlsx')  # открываем файл
    ws = wb.active  # открываем активную таблицу
    list_gup = []  # создаем список
    for row in ws.iter_rows(min_row=6, max_row=1077, min_col=6, max_col=9):  # перебираем строки
        row_data = [cell.value for cell in row]  # создаем список
        list_gup.append(row_data)  # добавляем в список
    return list_gup  # возвращаем список


def creation_contracts(row, formatted_date):
    doc = DocxTemplate("template/Шаблон_трудовой_договор.docx")

    context = {
        'name_surname': f" {row[1]} ",  # Ф.И.О. (Иванов Иван Иванович)
        'name_surname_completely': f" {row[2]} ",  # Ф.И.О. (Иванов И. И.)
        'date_admission': f" {formatted_date} ",  # Дата поступления
    }  # Должность
    doc.render(context)

    doc.save(f"готовые договора/{row[0]}_{row[2]}.docx")


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
        formatted_date = format_date(row[3])
        creation_contracts(row, formatted_date)
