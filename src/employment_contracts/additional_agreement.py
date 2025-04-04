# -*- coding: utf-8 -*-
from datetime import datetime

from loguru import logger

from src.database import read_from_db
from src.employment_contracts.filling_data import format_date
from src.employment_contracts.filling_in_the_data import generate_documents

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


if __name__ == "__main__":
    filling_ditional_agreement_health_reasons()
