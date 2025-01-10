# -*- coding: utf-8 -*-
from datetime import datetime

import openpyxl as op
from docxtpl import DocxTemplate
from loguru import logger


async def format_date(date):
    months = {
        1: "января",
        2: "февраля",
        3: "марта",
        4: "апреля",
        5: "мая",
        6: "июня",
        7: "июля",
        8: "августа",
        9: "сентября",
        10: "октября",
        11: "ноября",
        12: "декабря",
    }
    date = datetime.strptime(date, "%d.%m.%Y")
    return '" {:02d} " {} {} г.'.format(date.day, months[date.month], date.year)


async def open_list_gup():
    file = "../list_gup/Списочный_состав.xlsx"
    wb = op.load_workbook(file)  # открываем файл
    ws = wb.active  # открываем активную таблицу
    list_gup = []  # создаем список
    for row in ws.iter_rows(
            min_row=5, max_row=1100, min_col=0, max_col=34
    ):  # перебираем строки
        row_data = [cell.value for cell in row]  # создаем список
        list_gup.append(row_data)  # добавляем в список
    return list_gup  # возвращаем список


async def filling_data_hourly_rate(row, formatted_date, ending, file_dog):
    """Заполнение трудовых договоров. Часовая тарифная ставка"""

    doc = DocxTemplate(file_dog)
    date = row.a30

    # Проверяем, что дата не None и содержит три элемента после разделения
    if date is None or len(date.split(".")) != 3:
        return

    day, month, year = date.split(
        "."
    )  # Разделение даты, если в Excell файле стоит формат ячейки дата, то будет вызываться ошибка программы

    context = {
        "name_surname": f" {row.a5} ",  # Ф.И.О. (Иванов Иван Иванович)
        "name_surname_completely": f" {row.a6} ",  # Ф.И.О. (Иванов И. И.)
        "date_admission": f" {formatted_date} ",  # Дата поступления
        "ending": f"{ending}",  # Окончание ый или ая
        "post": f" {row.a3} ",  # Должность
        "district": f" {row.a1} ",  # Участок
        "salary": f" {row.a9} ",  # Часовая тарифная ставка или оклад
        "series_number": f"{row.a14}",  # Номер паспорта
        "phone": f"{row.a12}",  # Телефон
        "address": f"{row.a13}",  # Адрес
        "issue_date": f"{row.a15}",  # Дата выдачи
        "issued_by": f"{row.a16}",  # Кем выдан
        "code": f"{row.a17}",  # Код подразделения
        "official_salary": "часовая тарифная ставка",
        "official_salary_termination": "часовой тарифной ставки",
        "month_or_hour": "в час",
        "district_pro": f" {row.a19} ",  # Участок
        "employment_contract_number": f" {row.a25_номер_договора}",  # Номер трудового договора
        "day": f"{day}",  # День
        "month": f"{month}",  # Месяц
        "year": f"{year}",  # Год
        "graduation_from_profession": f" {row.a28} ",  # Профессия в родительном падеже
    }
    doc.render(context)
    doc.save(f"Готовые_договора/{row.a0}_{row.a4_табельный_номер}_{row.a5}.docx")


async def record_data_salary(row, formatted_date, ending, file_dog):
    """Заполнение трудовых договоров. Должностной оклад"""

    doc = DocxTemplate(file_dog)
    date = row.a30  # дата трудового договора

    # Проверяем, что дата не None и содержит три элемента после разделения
    if date is None or len(date.split(".")) != 3:
        return

    day, month, year = date.split(
        "."
    )  # Разделение даты, если в Excell файле стоит формат ячейки дата, то будет вызываться ошибка программы

    context = {
        "name_surname": f" {row.a5} ",  # Ф.И.О. (Иванов Иван Иванович)
        "name_surname_completely": f" {row.a6} ",  # Ф.И.О. (Иванов И. И.)
        "date_admission": f" {formatted_date} ",  # Дата поступления
        "ending": f"{ending}",  # Окончание ый или ая
        "post": f" {row.a3} ",  # Должность
        "district": f" {row.a1} ",  # Участок
        "salary": f" {row.a9} ",  # Часовая тарифная ставка или оклад
        "series_number": f"{row.a14}",  # Номер паспорта
        "phone": f"{row.a12}",  # Телефон
        "address": f"{row.a13}",  # Адрес
        "issue_date": f"{row.a15}",  # Дата выдачи
        "issued_by": f"{row.a16}",  # Кем выдан
        "code": f"{row.a17}",  # Код подразделения
        "official_salary": "должностной оклад",
        "official_salary_termination": "должностного оклада",
        "month_or_hour": "в месяц",
        "district_pro": f" {row.a19} ",  # Участок
        "employment_contract_number": f" {row.a25_номер_договора}",  # Номер трудового договора
        "day": f"{day}",  # День
        "month": f"{month}",  # Месяц
        "year": f"{year}",  # Год
        "graduation_from_profession": f" {row.a28} ",  # Профессия в родительном падеже
    }

    doc.render(context)
    doc.save(f"Готовые_договора/{row.a0}_{row.a4_табельный_номер}_{row.a5}.docx")


async def creation_contracts(row, formatted_date, ending):
    try:
        if row.a31 == "напечатанный":
            pass

        else:

            if float(row.a9) > 1000:  # Оклад
                if row.a34 == "None":
                    await record_data_salary(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор.docx",
                    )
                elif row.a34 == "Шаблон_трудовой_договор_уборщ_8_часов":  # 12 часов
                    await record_data_salary(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_уборщ_8_часов.docx",
                    )
                elif (
                        row.a34 == "Шаблон_трудовой_договор_8_часов_ИТР_подземные"
                ):  # 12 часов
                    await record_data_salary(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_8_часов_ИТР_подземные.docx",
                    )
                elif row.a34 == "Шаблон_трудовой_договор_12_часов":  # 12 часов
                    await record_data_salary(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_12_часов.docx",
                    )
                elif row.a34 == "Шаблон_трудовой_договор_6_часов":  # 6 часов
                    await record_data_salary(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_6_часов.docx",
                    )
                elif row.a34 == "Шаблон_трудовой_договор_7_часов":
                    await record_data_salary(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_7_часов.docx",
                    )
                elif (
                        row.a34
                        == "Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7"
                ):
                    await record_data_salary(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7.docx",
                    )
                elif row.a34 == "Шаблон_трудовой_договор_водителя_8_часов":
                    await record_data_salary(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_водителя_8_часов.docx",
                    )
                elif row.a34 == "Шаблон_трудовой_договор_8_часов_ИТР_без_вредности":
                    await record_data_salary(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_8_часов_ИТР_без_вредности.docx",
                    )

                elif row.a34 == "Шаблон_трудовой_договор_24_часа_без_вредн":
                    await record_data_salary(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_24_часа_без_вредн.docx",
                    )

            elif float(row.a9) < 1000:  # Часовая тарифная ставка
                if row.a34 == "None":
                    await filling_data_hourly_rate(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор.docx",
                    )
                elif row.a34 == "Шаблон_трудовой_договор_уборщ_8_часов":  # 12 часов
                    await filling_data_hourly_rate(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_уборщ_8_часов.docx",
                    )
                elif (
                        row.a34 == "Шаблон_трудовой_договор_8_часов_ИТР_подземные"
                ):  # 12 часов
                    await filling_data_hourly_rate(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_8_часов_ИТР_подземные.docx",
                    )
                elif row.a34 == "Шаблон_трудовой_договор_12_часов":  # 12 часов
                    await filling_data_hourly_rate(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_12_часов.docx",
                    )

                elif row.a34 == "ТД_6_час.раб.":  # 6 часов
                    await filling_data_hourly_rate(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/Рабочий/ТД_6_час.раб..docx",
                    )

                elif row.a34 == "Шаблон_трудовой_договор_7_часов":
                    await filling_data_hourly_rate(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_7_часов.docx",
                    )
                elif (
                        row.a34 == "Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7"
                ):
                    await filling_data_hourly_rate(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7.docx",
                    )
                elif row.a34 == "Шаблон_трудовой_договор_водителя_8_часов":
                    await filling_data_hourly_rate(
                        row,
                        formatted_date,
                        ending,
                        "Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_водителя_8_часов.docx",
                    )

    except Exception as e:
        logger.exception(e)


if __name__ == "__main__":
    open_list_gup()
