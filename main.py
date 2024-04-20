from docxtpl import DocxTemplate
import openpyxl as op


def open_list_gup():
    wb = op.load_workbook('list_gup/Списочный_состав.xlsx')  # открываем файл
    ws = wb.active  # открываем активную таблицу
    list_gup = []  # создаем список
    for row in ws.iter_rows(min_row=6, max_row=1077, min_col=6, max_col=8):  # перебираем строки
        row_data = [cell.value for cell in row]  # создаем список
        list_gup.append(row_data)  # добавляем в список
    return list_gup  # возвращаем список


def creation_contracts():
    doc = DocxTemplate("template/Шаблон_трудовой_договор.docx")

    context = {
        'name_surname': " Жабинский Виталий Викторович ",  # Ф.И.О.
        'name_surname_completely': " Жабинский В. В. ",  # Ф.И.О.
    }  # Должность
    doc.render(context)

    doc.save("шаблон-final.docx")


if __name__ == '__main__':
    # open_list_gup()
    # Пример использования функции
    parsed_data = open_list_gup()
    for row in parsed_data:
        print(row)

    # creation_contracts()
