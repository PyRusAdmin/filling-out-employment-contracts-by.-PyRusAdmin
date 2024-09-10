# -*- coding: utf-8 -*-
from datetime import datetime

import openpyxl as op
from docxtpl import DocxTemplate
from loguru import logger


def format_date(date):
    months = {
        1: "января", 2: "февраля", 3: "марта", 4: "апреля",
        5: "мая", 6: "июня", 7: "июля", 8: "августа",
        9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"
    }
    date = datetime.strptime(date, "%d.%m.%Y")
    return '" {:02d} " {} {} г.'.format(date.day, months[date.month], date.year)


def open_list_gup():
    file = 'list_gup/Списочный_состав.xlsx'
    wb = op.load_workbook(file)  # открываем файл
    ws = wb.active  # открываем активную таблицу
    list_gup = []  # создаем список
    for row in ws.iter_rows(min_row=5, max_row=1082, min_col=0, max_col=38):  # перебираем строки
        row_data = [cell.value for cell in row]  # создаем список
        list_gup.append(row_data)  # добавляем в список
    return list_gup  # возвращаем список


def filling_data_hourly_rate(row, formatted_date, ending, file_dog):
    """Часовая тарифная ставка"""

    doc = DocxTemplate(file_dog)
    date = row[34]
    day, month, year = date.split('.')  # Разделение даты, если в Excell файле стоит формат ячейки дата, то будет вызываться ошибка программы
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
        'employment_contract_number': f' {row[28]}',  # Номер трудового договора
        'day': f'{day}',  # День
        'month': f'{month}',  # Месяц
        'year': f'{year}',  # Год
        'graduation_from_profession': f' {row[32]} ',  # Профессия в родительном падеже
    }
    doc.render(context)
    doc.save(f"Готовые_договора/{row[0]}_{row[5]}_{row[6]}.docx")


def record_data_salary(row, formatted_date, ending, file_dog):
    """Должностной оклад"""

    doc = DocxTemplate(file_dog)
    date = row[34]
    day, month, year = date.split('.')  # Разделение даты, если в Excell файле стоит формат ячейки дата, то будет вызываться ошибка программы
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
        'employment_contract_number': f' {row[28]}',  # Номер трудового договора
        'day': f'{day}',  # День
        'month': f'{month}',  # Месяц
        'year': f'{year}',  # Год
        'graduation_from_profession': f' {row[32]} ',  # Профессия в родительном падеже
    }

    doc.render(context)
    doc.save(f"Готовые_договора/{row[0]}_{row[5]}_{row[6]}.docx")


def creation_contracts(row, formatted_date, ending):
    try:
        if row[35] == 'напечатанный':
            pass

        else:

            if row[11] > 1000: # Оклад
                if row[38] == 'None':
                    record_data_salary(row, formatted_date, ending,"Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор.docx")
                elif row[38] == 'Шаблон_трудовой_договор_уборщ_8_часов':  # 12 часов
                    record_data_salary(row, formatted_date, ending,"Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_уборщ_8_часов.docx")
                elif row[38] == 'Шаблон_трудовой_договор_8_часов_ИТР_подземные':  # 12 часов
                    record_data_salary(row, formatted_date, ending,"Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_8_часов_ИТР_подземные.docx")
                elif row[38] == 'Шаблон_трудовой_договор_12_часов':  # 12 часов
                    record_data_salary(row, formatted_date, ending,"Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_12_часов.docx")
                elif row[38] == 'Шаблон_трудовой_договор_6_часов':  # 6 часов
                    record_data_salary(row, formatted_date, ending,"Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_6_часов.docx")
                elif row[38] == 'Шаблон_трудовой_договор_7_часов':
                    record_data_salary(row, formatted_date, ending,"Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_7_часов.docx")
                elif row[38] == 'Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7':
                    record_data_salary(row, formatted_date, ending,"Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7.docx")
                elif row[38] == 'Шаблон_трудовой_договор_водителя_8_часов':
                    record_data_salary(row, formatted_date, ending,"Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_водителя_8_часов.docx")

            elif row[11] < 1000: # Часовая тарифная ставка
                if row[38] == 'None':
                    filling_data_hourly_rate(row, formatted_date, ending,"Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор.docx")
                elif row[38] == 'Шаблон_трудовой_договор_уборщ_8_часов':  # 12 часов
                    filling_data_hourly_rate(row, formatted_date, ending,"Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_уборщ_8_часов.docx")
                elif row[38] == 'Шаблон_трудовой_договор_8_часов_ИТР_подземные':  # 12 часов
                    filling_data_hourly_rate(row, formatted_date, ending,"Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_8_часов_ИТР_подземные.docx")
                elif row[38] == 'Шаблон_трудовой_договор_12_часов':  # 12 часов
                    filling_data_hourly_rate(row, formatted_date, ending,"Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_12_часов.docx")

                elif row[38] == 'ТД_6_час.раб.':  # 6 часов
                    filling_data_hourly_rate(row, formatted_date, ending,"Шаблоны_трудовых_договоров/Рабочий/ТД_6_час.раб..docx")

                elif row[38] == 'Шаблон_трудовой_договор_7_часов':
                    filling_data_hourly_rate(row, formatted_date, ending,"Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_7_часов.docx")
                elif row[38] == 'Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7':
                    filling_data_hourly_rate(row, formatted_date, ending,"Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7.docx")
                elif row[38] == 'Шаблон_трудовой_договор_водителя_8_часов':
                    filling_data_hourly_rate(row, formatted_date, ending,"Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_водителя_8_часов.docx")

    except Exception as e:
        logger.exception(e)




if __name__ == '__main__':
    open_list_gup()
