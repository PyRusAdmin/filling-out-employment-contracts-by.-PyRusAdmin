# -*- coding: utf-8 -*-
import os

import uvicorn
from fastapi import FastAPI
from fastapi import Form, Request, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from loguru import logger

from src.address_parsing import address_parsing
from src.checking_availability import get_missing_ids
from src.database import import_excel_to_db, database_cleaning_function
from src.filling_data import (
    formation_employment_contracts_filling_data,
    formation_and_filling_of_employment_contracts_for_idle_time_enterprise,
    formation_and_filling_of_part_time_employment_contracts,
    formation_and_filling_of_employment_contracts_for_transfer_to_another_job,
    filling_ditional_agreement_health_reasons, filling_notifications,
    filling_ditional_agreement_health_reasons_agreement_health
)
from src.formation_reduction_notification import formation_reduction_notification
from src.get import Employee
from src.parsing_comparison_file import parsing_document_1, compare_and_rewrite_professions
from src.receipt_contract import process_contracts_from_excel, process_single_contract  # ДОБАВЛЕНО

file = "data/list_gup/Списочный_состав.xlsx"

app = FastAPI()
# Монтируем статические файлы из папки "static"
app.mount("/static", StaticFiles(directory="static"), name="static")
# Монтируем папку data
app.mount("/data", StaticFiles(directory="data"), name="data")
templates = Jinja2Templates(directory="templates")
progress_messages = []  # список сообщений, которые будут отображаться в progress


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Главная страница"""
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/download_missing")
async def download_missing():
    if os.path.exists("missing.txt"):
        return FileResponse("missing.txt", filename="missing.txt", media_type="text/plain")
    return JSONResponse({"error": "missing.txt not found"}, status_code=404)


@app.get("/import_excel_form", response_class=HTMLResponse)
async def import_excel_form(request: Request):
    context = {
        "request": request,
        "url": "/data/list_gup/Списочный_состав.xlsx",  # Полный путь до файла
        "filename": "Списочный_состав.xlsx",  # Имя файла для загрузки
        "display_text": "Скачать список сотрудников"  # Отображаемый текст ссылки
    }
    return templates.TemplateResponse("import_excel_form.html", context)


@app.post("/import_excel")
async def import_excel(min_row: int = Form(...), max_row: int = Form(...)):
    """Импорт данных из файла"""
    try:
        logger.info(f"Запуск импорта данных с {min_row} по {max_row} строки.")

        await import_excel_to_db(min_row=min_row, max_row=max_row, file=file)

        return RedirectResponse(url="/", status_code=303)
    except Exception as e:
        logger.exception("Ошибка при импорте данных.")
        raise HTTPException(status_code=500, detail="Произошла ошибка при импорте данных.")


def search_employee_by_tab_number(tab_number):
    """Ищем данные сотрудника по табельному номеру"""
    try:
        return Employee.get(Employee.a4_табельный_номер == tab_number)
    except Employee.DoesNotExist:
        return None


@app.get("/get_contract", response_class=HTMLResponse)
async def get_contract_form(request: Request):
    """Страница получения данных сотрудника"""
    return templates.TemplateResponse("get_contract.html", {"request": request})


@app.post("/get_contract")
async def get_contract_process(
        tab_number: str = Form(...),
        request: Request = None
):
    """Обработка получения договора по табельному номеру"""
    try:
        logger.info(f"Запрос на получение договора для табельного номера: {tab_number}")

        # Проверяем валидность табельного номера
        if not tab_number.isdigit():
            logger.error(f"Некорректный табельный номер: {tab_number}")
            return templates.TemplateResponse(
                "get_contract.html",
                {
                    "request": request,
                    "error": "Табельный номер должен содержать только цифры"
                }
            )

        # Обрабатываем один договор
        result = process_single_contract(file, int(tab_number))

        if result is False:  # Сотрудник не найден
            return templates.TemplateResponse(
                "get_contract.html",
                {
                    "request": request,
                    "error": f"Сотрудник с табельным номером {tab_number} не найден"
                }
            )
        elif result is None:  # Договор уже напечатан
            return templates.TemplateResponse(
                "get_contract.html",
                {
                    "request": request,
                    "warning": f"Договор для табельного номера {tab_number} уже был напечатан"
                }
            )
        elif isinstance(result, str):  # Путь к файлу
            # Получаем имя файла из пути
            filename = os.path.basename(result)

            return templates.TemplateResponse(
                "get_contract.html",
                {
                    "request": request,
                    "success": f"Договор для табельного номера {tab_number} успешно создан!",
                    "download_url": f"/download_contract/{filename}",
                    "filename": filename
                }
            )
        else:
            return templates.TemplateResponse(
                "get_contract.html",
                {
                    "request": request,
                    "success": f"Договор для табельного номера {tab_number} успешно создан!"
                }
            )

    except Exception as e:
        logger.exception(f"Ошибка при обработке договора: {e}")
        return templates.TemplateResponse(
            "get_contract.html",
            {
                "request": request,
                "error": f"Произошла ошибка при создании договора: {str(e)}"
            }
        )


@app.get("/download_contract/{filename}")
async def download_contract(filename: str):
    """Скачивание созданного договора"""
    try:
        file_path = f"data/outgoing/Готовые_договора/{filename}"

        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="Файл не найден")

        return FileResponse(
            file_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        logger.exception(f"Ошибка при скачивании файла: {e}")
        raise HTTPException(status_code=500, detail="Ошибка при скачивании файла")


@app.get("/formation_employment_contracts", response_class=HTMLResponse)
async def formation_employment_contracts(request: Request):
    """Страница для формирования трудовых договоров"""
    return templates.TemplateResponse("formation_employment_contracts.html", {"request": request})


@app.get('/notification_compression', response_class=HTMLResponse)
async def notification_compression(request: Request):
    """Формирование уведомления о сокращении"""
    return templates.TemplateResponse("notification_compression.html", {"request": request})


@app.post("/action", response_class=HTMLResponse)
async def action(request: Request, user_input: str = Form(...)):
    """Выполнение действий"""
    logger.info(f"Выбранное действие: {user_input}")
    try:
        user_input = int(user_input)
        if user_input == 1:  # Парсинг данных из файла Excel
            await parsing_document_1(min_row=5, max_row=1084, column=5, column_1=8)
        elif user_input == 2:  # Формирование трудовых договоров
            await formation_employment_contracts_filling_data()
        elif user_input == 3:  # Сравнение и перезапись значений профессии в файле Excel счет начинается с 0
            await compare_and_rewrite_professions()
        elif user_input == 4:
            return RedirectResponse(url="/import_excel_form", status_code=303)

        elif user_input == 5:  # Получение договора
            return RedirectResponse(url="/get_contract", status_code=303)

        elif user_input == 6:  # Добавьте обработчик для выхода
            return RedirectResponse(url="/", status_code=303)
        elif user_input == 7:  # Очистка базы данных
            await database_cleaning_function(templates, request)
        elif user_input == 8:  # Заполнение договоров на простой
            await formation_and_filling_of_employment_contracts_for_idle_time_enterprise()
        elif user_input == 9:  # Заполнение договоров на не полную рабочую неделю
            await formation_and_filling_of_part_time_employment_contracts()
        elif user_input == 10:  # Дополнительное соглашение по состоянию здоровья
            await filling_ditional_agreement_health_reasons()
        elif user_input == 11:  # Дополнительное соглашение на перевод на другую должность (профессию)
            await formation_and_filling_of_employment_contracts_for_transfer_to_another_job()
        elif user_input == 12:  # Переход для формирования трудовых договоров и дополнительных соглашений
            return RedirectResponse(url="/formation_employment_contracts", status_code=303)
        elif user_input == 13:  # Переход для парсинга данных из файла
            await filling_notifications()  # Заполнение уведомлений для сотрудников
        elif user_input == 15:  # Формирование уведомление о сокращении
            await formation_reduction_notification()
        elif user_input == 17:  # Сверка уведомлений
            await get_missing_ids()

        elif user_input == 18:  # Парсинг адресов для конверта
            logger.info("Пользователь запустил (Парсинг адресов для конверта)")
            await address_parsing()
        elif user_input == 19:  # Формирование уведомления на сокращение
            return RedirectResponse(url="/notification_compression", status_code=303)

        elif user_input == 20:  # Заполнение дополнительного соглашения
            await filling_ditional_agreement_health_reasons_agreement_health()

        return RedirectResponse(url="/", status_code=303)
    except Exception as e:
        logger.exception(e)
        raise HTTPException(status_code=500, detail="Произошла ошибка.")


if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")