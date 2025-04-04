# -*- coding: utf-8 -*-
from loguru import logger

from src.employment_contracts.filling_in_the_data import generate_documents

# Табельные номера, для перевода на другую работу
transfer_to_another_job = [10711, 23495, ]


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
