import openpyxl as op
from docxtpl import DocxTemplate
from datetime import datetime
from loguru import logger


def get_all_data(file):
    """
    Получение всех данных из Excel-файла (строки 5–1115, колонки A–AI).

    Returns:
        list: список списков, где каждый список — строка таблицы
    """
    try:
        wb = op.load_workbook(file)
        ws = wb.active

        all_data = []

        # Перебираем строки с 5 по 1115
        for row_num in range(5, 1116):
            row_data = []
            # Собираем значения из колонок A–AI (1–35)
            for col_num in range(1, 36):
                row_data.append(ws.cell(row=row_num, column=col_num).value)

            # Добавляем только непустые строки (если хотя бы одна ячейка заполнена)
            if any(cell is not None for cell in row_data):
                all_data.append(row_data)

        return all_data

    except FileNotFoundError:
        logger.error(f"Файл не найден: {file}")
        return None
    except Exception as e:
        logger.error(f"Ошибка при работе с файлом: {e}")
        return None


def get_contract_by_number(search_value, all_data):
    """
    Поиск строки по табельному номеру в колонке E.

    Args:
        search_value: искомый табельный номер
        all_data: данные из Excel

    Returns:
        list or None: найденная строка или None
    """
    # Колонка E — это индекс 4 в списке (A=0, B=1, C=2, D=3, E=4)
    for row in all_data:
        if row[4] == search_value or str(row[4]) == str(search_value):
            return row

    return None


def format_date(date_str):
    """
    Форматирование даты из различных форматов

    Args:
        date_str: дата в виде строки или datetime объекта

    Returns:
        str: отформатированная дата
    """
    try:
        if isinstance(date_str, datetime):
            return date_str.strftime("%d.%m.%Y")
        elif isinstance(date_str, str):
            return date_str
        else:
            return "__.__.____"
    except Exception as e:
        logger.error(f"Ошибка форматирования даты: {e}")
        return "__.__.____"


def map_excel_row_to_dict(row_data):
    """
    Преобразование строки Excel в словарь с нужными полями

    Args:
        row_data: список значений из строки Excel

    Returns:
        dict: словарь с данными
    """
    # Маппинг колонок (индексы начинаются с 0)
    return {
        'a0': row_data[0] if len(row_data) > 0 else None,  # Колонка A
        'a1': row_data[1] if len(row_data) > 1 else None,  # Колонка B - Участок
        'a2': row_data[2] if len(row_data) > 2 else None,  # Колонка C
        'a3': row_data[3] if len(row_data) > 3 else None,  # Колонка D - Должность
        'a4_табельный_номер': row_data[4] if len(row_data) > 4 else None,  # Колонка E - Табельный номер
        'a5': row_data[5] if len(row_data) > 5 else None,  # Колонка F - ФИО полностью
        'a6': row_data[6] if len(row_data) > 6 else None,  # Колонка G - ФИО сокращенно
        'a7': row_data[7] if len(row_data) > 7 else None,  # Колонка H - Дата поступления
        'a8': row_data[8] if len(row_data) > 8 else None,  # Колонка I
        'a9': row_data[9] if len(row_data) > 9 else None,  # Колонка J - Оклад/ставка
        'a10': row_data[10] if len(row_data) > 10 else None,  # Колонка K
        'a11': row_data[11] if len(row_data) > 11 else None,  # Колонка L - Пол
        'a12': row_data[12] if len(row_data) > 12 else None,  # Колонка M - Телефон
        'a13': row_data[13] if len(row_data) > 13 else None,  # Колонка N - Адрес
        'a14': row_data[14] if len(row_data) > 14 else None,  # Колонка O - Серия и номер паспорта
        'a15': row_data[15] if len(row_data) > 15 else None,  # Колонка P - Дата выдачи
        'a16': row_data[16] if len(row_data) > 16 else None,  # Колонка Q - Кем выдан
        'a17': row_data[17] if len(row_data) > 17 else None,  # Колонка R - Код подразделения
        'a18': row_data[18] if len(row_data) > 18 else None,  # Колонка S
        'a19': row_data[19] if len(row_data) > 19 else None,  # Колонка T - Участок (pro)
        'a25_номер_договора': row_data[25] if len(row_data) > 25 else None,  # Колонка Z - Номер договора
        'a28': row_data[28] if len(row_data) > 28 else None,  # Колонка AC - Профессия
        'a30': row_data[30] if len(row_data) > 30 else None,  # Колонка AE - Дата договора
        'a31': row_data[31] if len(row_data) > 31 else None,  # Колонка AF - Статус (напечатанный)
        'a34': row_data[34] if len(row_data) > 34 else None,  # Колонка AI - Шаблон
    }


def generate_document(row_dict, file_dog, output_path):
    """
    Генерация документа из шаблона

    Args:
        row_dict: словарь с данными строки
        file_dog: путь к шаблону
        output_path: путь для сохранения
    """
    try:
        # Проверяем существование файла шаблона
        import os
        if not os.path.exists(file_dog):
            logger.error(f"Файл шаблона не найден: {file_dog}")
            return

        doc = DocxTemplate(file_dog)

        # Определяем окончание по полу
        ending = "ый" if row_dict.get('a11') == "Мужчина" else "ая"

        # Форматируем дату поступления
        formatted_date = format_date(row_dict.get('a7'))

        # Получение и проверка даты трудового договора
        date = row_dict.get('a30')

        if date is None or (isinstance(date, str) and len(date.split(".")) != 3):
            day, month, year = "--", "--", "----"
        else:
            date_str = str(date)
            parts = date_str.split(".")
            if len(parts) == 3:
                day, month, year = parts
            else:
                day, month, year = "--", "--", "----"

        # Подготовка контекста для заполнения
        context = {
            "name_surname": f" {row_dict.get('a5', '')} ",
            "name_surname_completely": f" {row_dict.get('a6', '')} ",
            "date_admission": f" {formatted_date} ",
            "ending": f"{ending}",
            "post": f" {row_dict.get('a3', '')} ",
            "district": f" {row_dict.get('a1', '')} ",
            "salary": f" {row_dict.get('a9', '')} ",
            "series_number": f"{row_dict.get('a14', '')}",
            "phone": f"{row_dict.get('a12', '')}",
            "address": f"{row_dict.get('a13', '')}",
            "issue_date": f"{row_dict.get('a15', '')}",
            "issued_by": f"{row_dict.get('a16', '')}",
            "code": f"{row_dict.get('a17', '')}",
            "official_salary": "должностной оклад",
            "official_salary_termination": "должностного оклада",
            "month_or_hour": "в месяц",
            "district_pro": f" {row_dict.get('a19', '')} ",
            "employment_contract_number": f" {row_dict.get('a25_номер_договора', '')}",
            "day": f"{day}",
            "month": f"{month}",
            "year": f"{year}",
            "graduation_from_profession": f" {row_dict.get('a28', '')} ",
        }

        doc.render(context)

        # Формирование имени файла
        filename = f"{row_dict.get('a0', 'unknown')}_{row_dict.get('a4_табельный_номер', 'unknown')}_{row_dict.get('a5', 'unknown')}.docx"
        full_path = f"{output_path}/{filename}"

        doc.save(full_path)
        logger.info(f"Документ сохранен: {full_path}")

    except Exception as e:
        logger.error(f"Ошибка генерации документа для {row_dict.get('a5', 'неизвестно')}: {e}")
        logger.error(f"Шаблон: {file_dog}")
        # Не пробрасываем исключение дальше, чтобы продолжить обработку других документов


def get_template_path(template_name, salary):
    """
    Определение пути к шаблону на основе названия и оклада

    Args:
        template_name: название шаблона из колонки AI
        salary: оклад/ставка

    Returns:
        str: путь к шаблону
    """
    base_path_itr = "data/docs_templates/Шаблоны_трудовых_договоров/ИТР"
    base_path_worker = "data/docs_templates/Шаблоны_трудовых_договоров/Рабочий"

    # Определяем базовый путь по окладу
    is_itr = float(salary) > 1000 if salary else False
    base_path = base_path_itr if is_itr else base_path_worker

    # Если шаблон не указан или None
    if template_name is None or template_name == "None":
        return f"{base_path}/Шаблон_трудовой_договор.docx"

    # Возвращаем полный путь к шаблону
    return f"{base_path}/{template_name}.docx"


def process_contracts_from_excel(excel_file, output_path="data/outgoing/Готовые_договора"):
    """
    Обработка всех трудовых договоров из Excel файла

    Args:
        excel_file: путь к Excel файлу
        output_path: путь для сохранения готовых договоров
    """
    start = datetime.now()
    logger.info(f"Время старта: {start}")

    # Загружаем все данные из Excel
    all_data = get_all_data(excel_file)

    if not all_data:
        logger.error("Не удалось загрузить данные из Excel")
        return

    logger.info(f"Загружено строк: {len(all_data)}")

    # Обрабатываем каждую строку
    processed_count = 0
    error_count = 0
    corrupted_templates = set()

    for row_data in all_data:
        row_dict = map_excel_row_to_dict(row_data)

        # Пропускаем уже напечатанные
        if row_dict.get('a31') == "напечатанный":
            continue

        # Пропускаем строки без оклада
        if not row_dict.get('a9'):
            continue

        try:
            salary = float(row_dict.get('a9'))
            template_name = row_dict.get('a34')

            # Получаем путь к шаблону
            template_path = get_template_path(template_name, salary)

            # Пропускаем, если шаблон уже помечен как поврежденный
            if template_path in corrupted_templates:
                logger.warning(f"Пропуск {row_dict.get('a5')} - шаблон поврежден: {template_path}")
                error_count += 1
                continue

            # Генерируем документ
            try:
                generate_document(row_dict, template_path, output_path)
                processed_count += 1
            except Exception as doc_error:
                if "Bad CRC-32" in str(doc_error):
                    corrupted_templates.add(template_path)
                    logger.error(f"Поврежденный шаблон: {template_path}")
                error_count += 1
                continue

        except Exception as e:
            logger.error(f"Ошибка обработки строки для {row_dict.get('a5', 'неизвестно')}: {e}")
            error_count += 1
            continue

    finish = datetime.now()
    logger.info(f"Время окончания: {finish}")
    logger.info(f"Время работы: {finish - start}")
    logger.info(f"Обработано договоров: {processed_count}")
    logger.info(f"Ошибок: {error_count}")

    if corrupted_templates:
        logger.warning("Поврежденные шаблоны:")
        for template in corrupted_templates:
            logger.warning(f"  - {template}")


def process_single_contract(excel_file, tabel_number, output_path="data/outgoing/Готовые_договора"):
    """
    Обработка одного трудового договора по табельному номеру

    Args:
        excel_file: путь к Excel файлу
        tabel_number: табельный номер сотрудника
        output_path: путь для сохранения готового договора

    Returns:
        str/bool/None: путь к файлу если успешно, False если не найден, None если уже напечатан
    """
    logger.info(f"Поиск сотрудника с табельным номером: {tabel_number}")

    # Загружаем все данные из Excel
    all_data = get_all_data(excel_file)

    if not all_data:
        logger.error("Не удалось загрузить данные из Excel")
        return False

    # Ищем строку с нужным табельным номером
    row_data = get_contract_by_number(tabel_number, all_data)

    if not row_data:
        logger.error(f"Сотрудник с табельным номером {tabel_number} не найден")
        return False

    row_dict = map_excel_row_to_dict(row_data)

    # Проверяем статус
    if row_dict.get('a31') == "напечатанный":
        logger.warning("Договор уже напечатан")
        return None

    try:
        salary = float(row_dict.get('a9'))
        template_name = row_dict.get('a34')

        # Получаем путь к шаблону
        template_path = get_template_path(template_name, salary)

        # Генерируем документ
        file_path = generate_document_with_return(row_dict, template_path, output_path)

        if file_path:
            logger.info(f"Договор успешно сгенерирован: {file_path}")
            return file_path
        else:
            return False

    except Exception as e:
        logger.exception(f"Ошибка обработки договора: {e}")
        return False


def generate_document_with_return(row_dict, file_dog, output_path):
    """
    Генерация документа из шаблона с возвратом пути к файлу

    Args:
        row_dict: словарь с данными строки
        file_dog: путь к шаблону
        output_path: путь для сохранения

    Returns:
        str or None: полный путь к созданному файлу или None при ошибке
    """
    try:
        # Проверяем существование файла шаблона
        import os
        if not os.path.exists(file_dog):
            logger.error(f"Файл шаблона не найден: {file_dog}")
            return None

        doc = DocxTemplate(file_dog)

        # Определяем окончание по полу
        ending = "ый" if row_dict.get('a11') == "Мужчина" else "ая"

        # Форматируем дату поступления
        formatted_date = format_date(row_dict.get('a7'))

        # Получение и проверка даты трудового договора
        date = row_dict.get('a30')

        if date is None or (isinstance(date, str) and len(date.split(".")) != 3):
            day, month, year = "--", "--", "----"
        else:
            date_str = str(date)
            parts = date_str.split(".")
            if len(parts) == 3:
                day, month, year = parts
            else:
                day, month, year = "--", "--", "----"

        # Подготовка контекста для заполнения
        context = {
            "name_surname": f" {row_dict.get('a5', '')} ",
            "name_surname_completely": f" {row_dict.get('a6', '')} ",
            "date_admission": f" {formatted_date} ",
            "ending": f"{ending}",
            "post": f" {row_dict.get('a3', '')} ",
            "district": f" {row_dict.get('a1', '')} ",
            "salary": f" {row_dict.get('a9', '')} ",
            "series_number": f"{row_dict.get('a14', '')}",
            "phone": f"{row_dict.get('a12', '')}",
            "address": f"{row_dict.get('a13', '')}",
            "issue_date": f"{row_dict.get('a15', '')}",
            "issued_by": f"{row_dict.get('a16', '')}",
            "code": f"{row_dict.get('a17', '')}",
            "official_salary": "должностной оклад",
            "official_salary_termination": "должностного оклада",
            "month_or_hour": "в месяц",
            "district_pro": f" {row_dict.get('a19', '')} ",
            "employment_contract_number": f" {row_dict.get('a25_номер_договора', '')}",
            "day": f"{day}",
            "month": f"{month}",
            "year": f"{year}",
            "graduation_from_profession": f" {row_dict.get('a28', '')} ",
        }

        doc.render(context)

        # Формирование имени файла
        filename = f"{row_dict.get('a0', 'unknown')}_{row_dict.get('a4_табельный_номер', 'unknown')}_{row_dict.get('a5', 'unknown')}.docx"
        full_path = f"{output_path}/{filename}"

        doc.save(full_path)
        logger.info(f"Документ сохранен: {full_path}")

        return full_path

    except Exception as e:
        logger.error(f"Ошибка генерации документа для {row_dict.get('a5', 'неизвестно')}: {e}")
        logger.error(f"Шаблон: {file_dog}")
        return None

# Пример использования
# if __name__ == "__main__":
#     excel_file = "data/list_gup/Списочный_состав.xlsx"

    # Вариант 1: Обработать все договоры
    # process_contracts_from_excel(excel_file)

    # Вариант 2: Обработать один договор по табельному номеру
    # process_single_contract(excel_file, 22148)