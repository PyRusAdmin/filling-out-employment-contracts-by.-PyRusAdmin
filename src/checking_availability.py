import json
import os
from loguru import logger


async def get_missing_ids():
    # Читаем разрешённые ID из data.json
    try:
        with open("data/data.json", "r", encoding="utf-8") as f:
            data = json.load(f)
        allowed_ids = set(data["ids"])
        logger.info(f"Загружено {len(allowed_ids)} разрешённых ID из data.json")
    except FileNotFoundError:
        logger.error("Файл data/data.json не найден!")
        return
    except KeyError:
        logger.error("В data.json отсутствует ключ 'ids'")
        return
    except json.JSONDecodeError as e:
        logger.error(f"Ошибка парсинга JSON: {e}")
        return

    folder = "output/Готовые_уведомления_сокращение"
    file_ids = {}

    # Собираем ID файлов и их имена
    for filename in os.listdir(folder):
        if filename.endswith(".docx"):
            parts = filename.split("_")
            if len(parts) > 1:
                try:
                    file_id = int(parts[1])
                    file_ids[file_id] = filename  # сохраняем имя файла по ID
                except ValueError:
                    logger.debug(f"Не удалось извлечь ID из файла: {filename}")
                    continue

    # Находим ID файлов, которых нет в data.json
    invalid_file_ids = [fid for fid in file_ids.keys() if fid not in allowed_ids]

    # Удаляем такие файлы
    for file_id in invalid_file_ids:
        filepath = os.path.join(folder, file_ids[file_id])
        try:
            os.remove(filepath)
            logger.warning(f"Удалён лишний файл: {file_ids[file_id]} (ID {file_id} не в data.json)")
        except Exception as e:
            logger.error(f"Не удалось удалить файл {file_ids[file_id]}: {e}")

    # Теперь находим missing_ids — ID из data.json, для которых нет файла
    missing_ids = sorted(allowed_ids - set(file_ids.keys()))

    # Сохраняем missing.txt
    if missing_ids:
        with open("missing.txt", "w", encoding="utf-8") as f:
            for mid in missing_ids:
                f.write(f"{mid}\n")
        logger.info(f"Создан файл missing.txt с {len(missing_ids)} отсутствующими ID")
    else:
        logger.info("Все ID из data.json имеют соответствующие файлы")

    logger.success("Проверка завершена")
