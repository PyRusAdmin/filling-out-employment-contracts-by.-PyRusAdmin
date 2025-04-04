# -*- coding: utf-8 -*-
from loguru import logger

from src.employment_contracts.filling_in_the_data import generate_documents

# на простой
not_a_full_work_week = [7123, 856, 1268, 1188, 5429, 23511, 4211, 3307, 10851, 10800, 3639, 11073, 8065, 13103,
                        13533, 7006, 7687, 15293, 17503, 17553, 12608, 600036, 12537, ]


async def creation_contracts_downtime(row, formatted_date, ending):
    try:
        # Проверяем, входит ли табельный номер в список
        if int(row.a4_табельный_номер) in not_a_full_work_week:
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
