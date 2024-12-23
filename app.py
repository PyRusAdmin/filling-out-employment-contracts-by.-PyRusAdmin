from datetime import datetime
from fastapi import FastAPI, Form, Request, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from loguru import logger
from database import import_excel_to_db, read_from_db
from filling_data import creation_contracts, format_date
from get import Employee
from parsing_comparison_file import parsing_document_1, compare_and_rewrite_professions
import asyncio

app = FastAPI()
templates = Jinja2Templates(directory="templates")
progress_messages = []  # список сообщений, которые будут отображаться в progress


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


async def run_import():
    # Импорт данных из Excel в базе данных в фоне
    await import_excel_to_db()


@app.get("/progress")
async def progress():
    async def generate():
        while True:
            if progress_messages:
                message = progress_messages.pop(0)
                yield f"data: {message}\n\n"
            await asyncio.sleep(1)

    return StreamingResponse(generate(), media_type="text/event-stream")


def search_employee_by_tab_number(tab_number):
    """Ищем данные сотрудника по табельному номеру"""
    try:
        # Поиск строки в базе данных по табельному номеру
        customer = Employee.get(Employee.a4_табельный_номер == tab_number)
        # Возвращаем строку со всеми данными
        return customer
    except Employee.DoesNotExist:
        # Возвращает None, если запись не найдена
        return None


@app.get("/get_contract", response_class=HTMLResponse)
async def get_contract_form(request: Request):
    return templates.TemplateResponse("get_contract.html", {"request": request})


@app.post("/get_contract")
async def get_contract(tab_number: str = Form(...)):
    logger.info(f"Введенный табельный номер: {tab_number}")  # Логируем введенный табельный номер
    if tab_number:
        data = search_employee_by_tab_number(tab_number)

        if data:
            return {
                "message": f"Данные для табельного номера {tab_number}",
                "data": {
                    "a0": data.a0,
                    "a1": data.a1,
                    # Добавьте остальные атрибуты по мере необходимости
                },
            }
        else:
            return {"message": f"Данные для табельного номера {tab_number} не найдены."}
    raise HTTPException(status_code=400, detail="Табельный номер не указан.")


@app.post("/action")
async def action(user_input: int = Form(...)):
    try:
        if user_input == 1:
            await parsing_document_1(min_row=5, max_row=1084, column=5, column_1=8)

        elif user_input == 2:
            start = datetime.now()  # фиксируем и выводим время старта работы кода
            logger.info("Время старта: " + str(start))
            data = await read_from_db()  # Считываем данные из базы данных
            # Выводим считанные данные на экран
            for row in data:
                logger.info(row)
                if row.a11 == "Мужчина":  # гендерное определение
                    ending = "ый"
                    await creation_contracts(
                        row,
                        await format_date(row.a7),  # дата поступления на предприятие
                        ending,
                    )
                elif row.a11 == "Женщина":  # гендерное определение
                    ending = "ая"
                    await creation_contracts(
                        row,
                        await format_date(row.a7),  # дата поступления на предприятие
                        ending,
                    )
            finish = datetime.now()  # фиксируем и выводим время окончания работы кода
            logger.info("Время окончания: " + str(finish))
            logger.info("Время работы: " + str(finish - start))  # вычитаем время старта из времени окончания

        elif user_input == 3:
            await compare_and_rewrite_professions()

        elif user_input == 4:
            # Запускаем асинхронный импорт данных и сразу отображаем страницу загрузки
            await import_excel_to_db()
            return RedirectResponse(url="/", status_code=303)

        elif user_input == 5:  # Получение трудового договора
            return RedirectResponse(url="/get_contract", status_code=303)

        elif user_input == 6:  # Выключение приложения
            exit()

        return RedirectResponse(url="/", status_code=303)
    except Exception as e:
        logger.exception(e)
        raise HTTPException(status_code=500, detail="Произошла ошибка.")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")