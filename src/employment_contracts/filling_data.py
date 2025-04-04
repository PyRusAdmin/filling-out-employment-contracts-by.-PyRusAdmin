# -*- coding: utf-8 -*-
from datetime import datetime

import openpyxl as op
from loguru import logger

from src.employment_contracts.filling_in_the_data import generate_documents


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
    file = "data/list_gup/Списочный_состав.xlsx"
    wb = op.load_workbook(file)  # открываем файл
    ws = wb.active  # открываем активную таблицу
    list_gup = []  # создаем список
    for row in ws.iter_rows(
            min_row=5, max_row=1112, min_col=0, max_col=34
    ):  # перебираем строки
        row_data = [cell.value for cell in row]  # создаем список
        list_gup.append(row_data)  # добавляем в список
    return list_gup  # возвращаем список


async def creation_contracts(row, formatted_date, ending):
    try:
        if row.a31 == "напечатанный":
            pass

        else:

            if float(row.a9) > 1000:  # Оклад
                if row.a34 == "None":
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_уборщ_8_часов":  # 12 часов
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_уборщ_8_часов.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif (
                        row.a34 == "Шаблон_трудовой_договор_8_часов_ИТР_подземные"
                ):  # 12 часов
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_8_часов_ИТР_подземные.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_12_часов":  # 12 часов
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_12_часов.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_6_часов":  # 6 часов
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_6_часов.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_7_часов":
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_7_часов.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7":
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_водителя_8_часов":
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_водителя_8_часов.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_8_часов_ИТР_без_вредности":
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/ИТР/Шаблон_трудовой_договор_8_часов_ИТР_без_вредности.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )

                elif row.a34 == "Шаблон_трудовой_договор_24_часа_без_вредн":
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_24_часа_без_вредн.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )

            elif float(row.a9) < 1000:  # Часовая тарифная ставка
                if row.a34 == "None":
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_уборщ_8_часов":  # 12 часов
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contractsШаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_уборщ_8_часов.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_8_часов_ИТР_подземные":  # 12 часов
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contractsШаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_8_часов_ИТР_подземные.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_12_часов":  # 12 часов
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_12_часов.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )

                elif row.a34 == "ТД_6_час.раб.":  # 6 часов
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/Рабочий/ТД_6_час.раб..docx",
                        output_path="data/outgoing/Готовые_договора"
                    )

                elif row.a34 == "Шаблон_трудовой_договор_7_часов":
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_7_часов.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7":
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contractsШаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_8_часов_ИТР_контора_вредность_не_норм_7.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )
                elif row.a34 == "Шаблон_трудовой_договор_водителя_8_часов":
                    await generate_documents(
                        row=row,
                        formatted_date=formatted_date,
                        ending=ending,
                        file_dog="data/templates_contracts/Шаблоны_трудовых_договоров/Рабочий/Шаблон_трудовой_договор_водителя_8_часов.docx",
                        output_path="data/outgoing/Готовые_договора"
                    )

    except Exception as e:
        logger.exception(e)


if __name__ == "__main__":
    open_list_gup()
