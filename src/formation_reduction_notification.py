from datetime import datetime

from loguru import logger

from src.database import read_from_db
from src.filling_data import generate_documents, format_date


async def formation_reduction_notification():
    logger.info("Пользователь выбрал формирование уведомление о сокращении")

    """Заполнение уведомлений"""

    start = datetime.now()
    logger.info(f"Время старта: {start}")
    data = await read_from_db()
    for row in data:
        logger.info(row)
        ending = "ый" if row.a11 == "Мужчина" else "ая"

        await generate_documents(
            row=row,
            formatted_date=await format_date(row.a7),
            ending=ending,
            file_dog="data/templates_contracts/Сокращение/уведомления.docx",  # шаблон уведомления
            output_path="data/outgoing/Готовые_уведомления_сокращение"  # папка для сохранения уведомления
        )

    finish = datetime.now()
    logger.info(f"Время окончания: {finish}\n\nВремя работы: {finish - start}")
