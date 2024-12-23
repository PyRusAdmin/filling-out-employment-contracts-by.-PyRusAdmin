# -*- coding: utf-8 -*-
from docxtpl import DocxTemplate
from loguru import logger

not_a_full_work_week = [7123, 856, 1268, 1188, 5429, 23511, 4211, 3307, 10851, 10800, 3639, 11073, 8065, 13103,
                        13533, 7006, 7687, 15293, 17503, 17553, 12608, 600036, 12537, ]


async def record_data_salary_downtime(row, formatted_date, ending, file_dog):
    """Должностной оклад"""

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
    doc.save(f"Готовые_дополнительные_договора/{row.a0}_{row.a4_табельный_номер}_{row.a5}.docx")


async def creation_contracts_downtime(row, formatted_date, ending):
    try:
        # Проверяем, входит ли табельный номер в список
        if int(row.a4_табельный_номер) in not_a_full_work_week:
            await record_data_salary_downtime(
                row,
                formatted_date,
                ending,
                "Шаблоны_дополнительных_соглашений/доп_соглашение_к_трудовому_договору_простой.docx",
            )
        else:
            logger.info(f"Табельный номер {row.a4_табельный_номер} не входит в список. Договор не будет сформирован.")
    except Exception as e:
        logger.exception(f"Ошибка при формировании договора для табельного номера {row.a4_табельный_номер}: {e}")
