# -*- coding: utf-8 -*-
from datetime import datetime

import openpyxl as op
from docxtpl import DocxTemplate
from loguru import logger

from src.database import read_from_db


async def generate_documents(row, formatted_date, ending, file_dog, output_path):
    doc = DocxTemplate(file_dog)  # Загрузка шаблона
    # Получение и проверка даты трудового договора
    date = row.a30  # дата трудового договора
    # Проверяем, что дата не None и содержит три элемента после разделения
    if date is None or len(date.split(".")) != 3:
        return

    day, month, year = date.split(".")  # Разделение даты на компоненты
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

    doc.render(context)  # Рендеринг документа
    # Формирование имени файла
    filename = f"{row.a0}_{row.a4_табельный_номер}_{row.a5}.docx"
    full_path = f"{output_path}/{filename}"

    doc.save(full_path)  # Сохранение документа


# Заполнение уведомлений
async def filling_notifications():
    """Заполнение уведомлений"""
    start = datetime.now()
    logger.info(f"Время старта: {start}")
    data = await read_from_db()
    for row in data:
        logger.info(row)
        ending = "ый" if row.a11 == "Мужчина" else "ая"
        # await creation_contracts(row, await format_date(row.a7), ending)

        await generate_documents(
            row=row,
            formatted_date=await format_date(row.a7),
            ending=ending,
            file_dog="data/templates_contracts/уведомления/уведомление.docx",
            output_path="data/outgoing/Готовые_уведомления"
        )

    finish = datetime.now()
    logger.info(f"Время окончания: {finish}\n\nВремя работы: {finish - start}")


# на простой
not_a_full_work_weeks = [7123, 856, 1268, 1188, 5429, 23511, 4211, 3307, 10851, 10800, 3639, 11073, 8065, 13103,
                         13533, 7006, 7687, 15293, 17503, 17553, 12608, 600036, 12537, ]


async def creation_contracts_downtime(row, formatted_date, ending):
    try:
        # Проверяем, входит ли табельный номер в список
        if int(row.a4_табельный_номер) in not_a_full_work_weeks:
            await generate_documents(
                row=row,
                formatted_date=formatted_date,
                ending=ending,
                file_dog="data/templates_contracts/Шаблоны_доп_соглашений/доп_соглашение_к_труд_дог_простой.docx",
                output_path="data/outgoing/Готовые_дополнительные_договора"
            )
        else:
            logger.info(f"Табельный номер {row.a4_табельный_номер} не входит в список. Договор не будет сформирован.")
    except Exception as e:
        logger.exception(f"Ошибка при формировании договора для табельного номера {row.a4_табельный_номер}: {e}")


# не полная рабочая неделя
not_a_full_work_week = [7123, 12212, 856, 1268, 1188, 5429, 23173, 23511, 4211, 3307, 10851, 10800, 3639, 11073, 8065,
                        13103, 13533, 7006, 7687, 15293, 17503, 17553, 12608, 600036, 12537, 23492, ]


async def creation_contracts_downtime_week(row, formatted_date, ending):
    try:
        # Проверяем, входит ли табельный номер в список
        if int(row.a4_табельный_номер) in not_a_full_work_week:
            await generate_documents(
                row=row,
                formatted_date=formatted_date,
                ending=ending,
                file_dog="data/templates_contracts/Шаблоны_доп_соглашений/доп_соглашение_к_труд_дог_неп_раб_время.docx",
                output_path="data/outgoing/Готовые_дополнительные_соглашения_не_полная_рабочая_неделя"
            )
        else:
            logger.info(f"Табельный номер {row.a4_табельный_номер} не входит в список. Договор не будет сформирован.")
    except Exception as e:
        logger.exception(f"Ошибка при формировании договора для табельного номера {row.a4_табельный_номер}: {e}")


# Табельные номера, для перевода на другую работу
transfer_to_another_job = [10711, 23495, 15675]


async def creation_contracts_another_job(row, formatted_date, ending):
    try:
        # Проверяем, входит ли табельный номер в список
        if int(row.a4_табельный_номер) in transfer_to_another_job:
            await generate_documents(
                row=row,
                formatted_date=formatted_date,
                ending=ending,
                file_dog="data/templates_contracts/Шаблоны_доп_соглашений/доп_соглашение_к_труд_дог_перевод.docx",
                output_path="data/outgoing/Готовые_дополнительные_соглашения_перевод_на_другую_работу"
            )
        else:
            logger.info(f"Табельный номер {row.a4_табельный_номер} не входит в список. Договор не будет сформирован.")
    except Exception as e:
        logger.exception(f"Ошибка при формировании договора для табельного номера {row.a4_табельный_номер}: {e}")


# Дополнительное соглашение для увольнения
additional_agreement_list = [23173]


async def creation_contracts_additional_agreement(row, formatted_date, ending):
    try:
        # Проверяем, входит ли табельный номер в список
        if int(row.a4_табельный_номер) in additional_agreement_list:
            await generate_documents(
                row=row,
                formatted_date=formatted_date,
                ending=ending,
                file_dog="data/templates_contracts/договоры_компенсации/расторжение_ЗД.docx",
                output_path="data/outgoing/доп_согл_нпн"
            )
        else:
            logger.info(f"Табельный номер {row.a4_табельный_номер} не входит в список. Договор не будет сформирован.")
    except Exception as e:
        logger.exception(f"Ошибка при формировании договора для табельного номера {row.a4_табельный_номер}: {e}")


async def filling_ditional_agreement_health_reasons():
    """Заполнение дополнительного соглашения по состоянию здоровья"""
    start = datetime.now()
    logger.info(f"Время старта: {start}")
    data = await read_from_db()
    for row in data:
        logger.info(row)
        ending = "ый" if row.a11 == "Мужчина" else "ая"
        await creation_contracts_additional_agreement(row, await format_date(row.a7), ending)
    finish = datetime.now()
    logger.info(f"Время окончания: {finish}")
    logger.info(f"Время работы: {finish - start}")


async def formation_and_filling_of_employment_contracts_for_transfer_to_another_job():
    """Формирование трудовых договоров на переход на другую работу"""
    start = datetime.now()
    logger.info(f"Время старта: {start}")
    data = await read_from_db()
    for row in data:
        logger.info(row)
        ending = "ый" if row.a11 == "Мужчина" else "ая"
        await creation_contracts_another_job(row, await format_date(row.a7), ending)
    finish = datetime.now()
    logger.info(f"Время окончания: {finish}\n\nВремя работы: {finish - start}")


async def formation_and_filling_of_part_time_employment_contracts():
    """Формирование трудовых договоров на не полную рабочую неделю"""
    start = datetime.now()
    logger.info(f"Время старта: {start}")
    data = await read_from_db()
    for row in data:
        logger.info(row)
        ending = "ый" if row.a11 == "Мужчина" else "ая"
        await creation_contracts_downtime_week(row, await format_date(row.a7), ending)
    finish = datetime.now()
    logger.info(f"Время окончания: {finish}\n\nВремя работы: {finish - start}")


async def formation_and_filling_of_employment_contracts_for_idle_time_enterprise():
    """Формирование трудовых договоров на простой предприятия"""
    start = datetime.now()
    logger.info(f"Время старта: {start}")
    data = await read_from_db()
    for row in data:
        logger.info(row)
        ending = "ый" if row.a11 == "Мужчина" else "ая"
        await creation_contracts_downtime(row, await format_date(row.a7), ending)
    finish = datetime.now()
    logger.info(f"Время окончания: {finish}\n\nВремя работы: {finish - start}")


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


async def formation_employment_contracts_filling_data():
    """Формирование трудовых договоров"""
    start = datetime.now()
    logger.info(f"Время старта: {start}")
    data = await read_from_db()
    for row in data:
        logger.info(row)
        ending = "ый" if row.a11 == "Мужчина" else "ая"
        await creation_contracts(row, await format_date(row.a7), ending)
    finish = datetime.now()
    logger.info(f"Время окончания: {finish}\n\nВремя работы: {finish - start}")


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


