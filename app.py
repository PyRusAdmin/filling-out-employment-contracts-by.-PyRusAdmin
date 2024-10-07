import asyncio
from datetime import datetime

from loguru import logger
from quart import Quart, render_template, request, redirect, url_for

from database import import_excel_to_db, read_from_db
from filling_data import creation_contracts, format_date
from get import Employee
from parsing_comparison_file import parsing_document_1, compare_and_rewrite_professions

app = Quart(__name__)

progress_messages = []  # список сообщений, которые будут отображаться в progress


@app.route("/")
async def index():
    return await render_template("index.html")


# @app.route("/loading")
# async def loading():
#     """Сообщение, что база данных формируется"""
#     return await render_template("loading.html")


async def run_import():
    # Импорт данных из Excel в базе данных в фоне
    await import_excel_to_db()


@app.route("/progress")
async def progress():
    async def generate():
        while True:
            if progress_messages:
                message = progress_messages.pop(0)
                yield f"data: {message}\n\n"
            await asyncio.sleep(1)

    return generate()


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


@app.route("/get_contract", methods=["GET", "POST"])
async def get_contract():
    if request.method == "POST":
        tab_number = (await request.form).get("tab_number")
        logger.info(
            f"Введенный табельный номер: {tab_number}"
        )  # Логируем введенный табельный номер
        if tab_number:
            data = search_employee_by_tab_number(tab_number)

            if data:
                return f"Данные для табельного номера {tab_number}: {data.a0, data.a1, data.a2, data.a3,
                data.a4_табельный_номер, data.a5, data.a6, data.a7, data.a8, data.a9, data.a10, data.a11, data.a12,
                data.a13, data.a14, data.a15, data.a16, data.a17, data.a18, data.a19, data.a20, data.a21, data.a22,
                data.a23, data.a24, data.a25_номер_договора, data.a26, data.a27, data.a28, data.a29, data.a30,
                data.a31, data.a32, data.a33, data.a34}"
            else:
                return f"Данные для табельного номера {tab_number} не найдены."
        return "Табельный номер не указан."
    return await render_template(
        "get_contract.html"
    )  # Отображаем форму для GET-запроса


@app.route("/action", methods=["POST"])
async def action():
    try:
        user_input = int((await request.form)["user_input"])

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
                    # logger.info(row.some_attribute)  # выводим данные в консоль
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
            logger.info(
                "Время работы: " + str(finish - start)
            )  # вычитаем время старта из времени окончания

        elif user_input == 3:
            await compare_and_rewrite_professions()

        elif user_input == 4:
            # Запускаем асинхронный импорт данных и сразу отображаем страницу загрузки
            await import_excel_to_db()
            return redirect(url_for("index"))

        elif user_input == 5:  # Получение трудового договора
            return redirect(
                url_for("get_contract")
            )  # Перенаправляем на страницу get_contract

        elif user_input == 6:  # Выключение приложения
            exit()

        return redirect(url_for("index"))
    except Exception as e:
        logger.exception(e)


if __name__ == "__main__":
    app.run(debug=True)
