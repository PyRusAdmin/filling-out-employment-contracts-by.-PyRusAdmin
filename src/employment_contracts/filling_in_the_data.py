# -*- coding: utf-8 -*-

from docxtpl import DocxTemplate


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
