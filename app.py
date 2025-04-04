from datetime import datetime

from fastapi import FastAPI, Form, Request, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from loguru import logger

from src.database import import_excel_to_db, read_from_db, clear_database
from src.employment_contracts.additional_agreement import creation_contracts_additional_agreement
from src.employment_contracts.filling_a_shortened_work_week import creation_contracts_downtime_week
from src.employment_contracts.filling_data import creation_contracts, format_date
from src.employment_contracts.filling_plant_downtime import creation_contracts_downtime
from src.get import Employee
from src.parsing_comparison_file import parsing_document_1, compare_and_rewrite_professions

app = FastAPI()
templates = Jinja2Templates(directory="templates")
progress_messages = []  # список сообщений, которые будут отображаться в progress


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Главная страница"""
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/import_excel_form", response_class=HTMLResponse)
async def import_excel_form(request: Request):
    """Страница импорта данных из файла"""
    return templates.TemplateResponse("import_excel_form.html", {"request": request})


@app.post("/import_excel")
async def import_excel(min_row: int = Form(...), max_row: int = Form(...)):
    """Импорт данных из файла"""
    try:
        logger.info(f"Запуск импорта данных с {min_row} по {max_row} строки.")
        await import_excel_to_db(min_row=min_row, max_row=max_row)
        return RedirectResponse(url="/", status_code=303)
    except Exception as e:
        logger.exception("Ошибка при импорте данных.")
        raise HTTPException(status_code=500, detail="Произошла ошибка при импорте данных.")


def search_employee_by_tab_number(tab_number):
    """Ищем данные сотрудника по табельному номеру"""
    try:
        customer = Employee.get(Employee.a4_табельный_номер == tab_number)
        return customer
    except Employee.DoesNotExist:
        return None


@app.get("/get_contract", response_class=HTMLResponse)
async def get_contract_form(request: Request):
    """Страница получения данных сотрудника"""
    return templates.TemplateResponse("get_contract.html", {"request": request})


@app.post("/get_contract", response_class=HTMLResponse)
async def get_contract(request: Request, tab_number: str = Form(...), ):
    """Получение данных сотрудника"""
    logger.info(f"Введенный табельный номер: {tab_number}")

    if tab_number:
        data = search_employee_by_tab_number(tab_number)
        if data:
            return templates.TemplateResponse("contract_data.html", {
                "request": request,
                "tab_number": tab_number,
                "data": {
                    'КСП': data.a0,
                    'наименование ксп': data.a1,
                    'Категория': data.a2,
                    'профессия': data.a3,
                    'Таб №': data.a4_табельный_номер,
                    'Ф.И.О.': data.a5,
                    'Ф.И.О. (сокращенно)': data.a6,
                    'Дата приема': data.a7,
                    'Дата увольнения': data.a8,
                    'Тариф / Оклад': data.a9,
                    'Дата рождения': data.a10,
                    'ПОЛ': data.a11,
                    'Телефон': data.a12,
                    'Адрес': data.a13,
                    'Серия код': data.a14,
                    'Дата выдачи': data.a15,
                    'Кем выдан': data.a16,
                    'Код подразделения': data.a17,
                    'Продолжительность рабочего дня': data.a18,
                    'Окончание': data.a19,
                    'За ненорм.': data.a20,
                    'Особый характер труда': data.a21,
                    'За вредные условия труда': data.a22,
                    'Начальники': data.a23,
                    'Статус': data.a24,
                    'Номер договора': data.a25_номер_договора,
                    'Профессия': data.a26,
                    'Профессия с разрядами': data.a27,
                    'Профессия в родительном падеже': data.a28,
                    'Дополнительный отпуск': data.a29,
                    'дата договора': data.a30,
                    'Готовность': data.a31,
                    'Дата перевода (приема) и номер приказа': data.a32,
                    'Договор / дополнительное соглашение': data.a33,
                    'Тип шаблона': data.a34
                }
            })
        else:
            return {"message": f"Данные для табельного номера {tab_number} не найдены."}
    raise HTTPException(status_code=400, detail="Табельный номер не указан.")


@app.post("/action", response_class=HTMLResponse)
async def action(request: Request, user_input: str = Form(...)):
    """Выполнение действий"""
    logger.info(f"Выбранное действие: {user_input}")
    try:
        user_input = int(user_input)

        if user_input == 1:
            await parsing_document_1(min_row=5, max_row=1084, column=5, column_1=8)

        elif user_input == 2:
            start = datetime.now()
            logger.info(f"Время старта: {start}")
            data = await read_from_db()
            for row in data:
                logger.info(row)
                ending = "ый" if row.a11 == "Мужчина" else "ая"
                await creation_contracts(row, await format_date(row.a7), ending)
            finish = datetime.now()
            logger.info(f"Время окончания: {finish}")
            logger.info(f"Время работы: {finish - start}")

        elif user_input == 3:
            await compare_and_rewrite_professions()


        elif user_input == 4:

            return RedirectResponse(url="/import_excel_form", status_code=303)


        elif user_input == 5:
            return RedirectResponse(url="/get_contract", status_code=303)


        elif user_input == 7:
            try:
                await clear_database()
                logger.info("База данных успешно очищена.")
                return templates.TemplateResponse("database_cleanup.html",
                                                  {"request": request, "message": "База данных успешно очищена!"})
            except Exception as e:
                logger.exception("Ошибка при очистке базы данных.")
                return templates.TemplateResponse("database_cleanup.html",
                                                  {"request": request, "message": f"Ошибка: {e}"})

        elif user_input == 8:
            start = datetime.now()
            logger.info(f"Время старта: {start}")
            data = await read_from_db()
            for row in data:
                logger.info(row)
                ending = "ый" if row.a11 == "Мужчина" else "ая"
                await creation_contracts_downtime(row, await format_date(row.a7), ending)
            finish = datetime.now()
            logger.info(f"Время окончания: {finish}")
            logger.info(f"Время работы: {finish - start}")

        elif user_input == 9:
            start = datetime.now()
            logger.info(f"Время старта: {start}")
            data = await read_from_db()
            for row in data:
                logger.info(row)
                ending = "ый" if row.a11 == "Мужчина" else "ая"
                await creation_contracts_downtime_week(row, await format_date(row.a7), ending)
            finish = datetime.now()
            logger.info(f"Время окончания: {finish}")
            logger.info(f"Время работы: {finish - start}")

        elif user_input == 10:
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

        return RedirectResponse(url="/", status_code=303)

    except Exception as e:
        logger.exception(e)
        raise HTTPException(status_code=500, detail="Произошла ошибка.")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")
