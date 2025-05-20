import os
from io import BytesIO
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse
from fastapi.templating import Jinja2Templates
from dotenv import load_dotenv
import logging
import uvicorn
import pandas as pd
import requests

# from config import URL, id, INCOMING_URL, FILE_NAME

# Логирование
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Инициализация FastAPI-приложения
app = FastAPI()

# Конфигурация Jinja2 для рендеринга HTML
templates = Jinja2Templates(directory="templates")

load_dotenv()

# Используем переменные окружения
URL = os.getenv("URL")
INCOMING_URL = os.getenv("INCOMING_URL")
FILE_NAME = os.getenv("FILE_NAME")
ID = os.getenv("ID")


# Функция для получения ссылки на файл из Bitrix24
def get_children(folder_id):
    params = {
        'id': folder_id,
    }
    response = requests.get(URL, params=params)
    if response.status_code == 200:
        return response.json().get('result', {})
    else:
        raise Exception(f'Ошибка при получении ссылки на файл: {response.status_code}')


# Функция для загрузки файла из Bitrix24
def get_word_file(file_name: str) -> BytesIO:
    """
    Получает файл из Bitrix24 по имени и возвращает его в виде BytesIO.
    """
    NDT_FOLDERS = ID  # ID папки в Bitrix24
    df_file_in_folder = pd.DataFrame(get_children(NDT_FOLDERS))
    try:
        download_url = df_file_in_folder[
            df_file_in_folder['NAME'] == file_name
            ]['DOWNLOAD_URL'].values[0]
        return download_file_from_bitrix(download_url)
    except IndexError:
        raise Exception(f'Файл с именем "{file_name}" не найден в папке.')


# Функция для скачивания файла из Bitrix24
def download_file_from_bitrix(download_url: str) -> BytesIO:
    """
    Скачивает файл по URL и возвращает его в виде BytesIO.
    """
    response = requests.get(download_url)
    if response.status_code == 200:
        return BytesIO(response.content)
    else:
        raise Exception(f'Ошибка при скачивании файла: {response.status_code}')


# Функция для загрузки данных из Excel (из Bitrix24)
def load_excel_data(file_name):
    """
    Данная функция загружает данные с определенными параметрами (target_sheets - названиями листов)
    :param file_name: Название документа.
    :return: Список / таблицу с нужными листами в документе.
    """
    try:
        # Загружаем файл из Bitrix24
        file_content = get_word_file(file_name)
        data = []
        # Список листов, которые нужно обработать
        target_sheets = ["Номера", "Изменение материалов"]

        # Читаем Excel-файл с помощью pandas
        with pd.ExcelFile(file_content) as xls:
            for sheet_name in xls.sheet_names:
                if sheet_name not in target_sheets:
                    continue  # Пропускаем листы, которые не входят в список

                logger.info(f"Загружаем данные с листа: {sheet_name}")
                df = pd.read_excel(xls, sheet_name=sheet_name, header=3)  # Учитываем заголовок на 4-й строке

                # Заменяем NaN на пустые строки
                df = df.fillna("")

                for _, row in df.iterrows():
                    # Добавляем название листа к каждой строке
                    data.append((sheet_name, tuple(row)))
        return data
    except Exception as e:
        logger.error(f"Ошибка при загрузке Excel-файла: {e}")
        return []


# Глобальная переменная для хранения данных из Excel
excel_data = load_excel_data(FILE_NAME)


# if not excel_data:
#     logger.error("Не удалось загрузить данные из Excel-файла.")
# else:
#     logger.info(f"Загружено {len(excel_data)} строк из Excel-файла.")


# Функция для разбиения длинного сообщения
def split_message(message, max_length=4096):
    """
        Разбивает длинное сообщение на части, каждая из которых не превышает max_length символов.

        :param message: Исходное сообщение.
        :param max_length: Максимальная длина одной части сообщения.
        :return: Список частей сообщения.
    """
    return [message[i:i + max_length] for i in range(0, len(message), max_length)]


# Функция для форматирования строки из Excel
def format_excel_row(sheet_name, row):
    """
        Форматирует строку из Excel-файла для вывода в читаемом виде.

        :param sheet_name: Название листа Excel.
        :param row: Кортеж данных строки из Excel.
        :return: Отформатированная строка.
    """
    try:
        id_value = str(row[0]) if len(row) > 0 and row[0] is not None else ""
        name_value = str(row[1]) if len(row) > 1 and row[1] is not None else ""
        name_content = str(row[2]) if len(row) > 1 and row[2] is not None else ""
        wagon_info = str(row[4]) if len(row) > 4 and row[4] is not None else ""
        link = str(row[7]) if len(row) > 7 and row[7] is not None else ""
        units_of_measurement = str(row[13]) if len(row) > 13 and row[13] is not None else ""

        # Формируем строку с указанием листа
        if sheet_name == "Изменение материалов":
            formatted_string = f'{wagon_info}'
            return formatted_string
        else:
            units_part = f", _ {units_of_measurement}" if units_of_measurement else ""
            formatted_string = (
                f"Информация по материалу:\n\n{name_value}\n{name_content}\n\n"
                f"Заказные номера:\n{wagon_info}\n\n"
                f"Ласточка / Финист:\n[{id_value}] ({name_value}{units_part})\n\n"
                f"Сапсан:\n[{id_value}] ({wagon_info}, {name_value}{units_part})\n\n"
            )
            if link:
                formatted_string += f"\n\nПримечание:\n{link}"
            return formatted_string
    except Exception as e:
        logger.error(f"Ошибка при форматировании строки: {e}")
        return ", ".join(str(cell) for cell in row)


# @app.get("/", response_class=HTMLResponse)
# async def flask_render():
#     """
#         Обрабатывает GET-запрос для отображения данных из Excel-файла в виде HTML-таблицы.
#
#         :return: HTML-страница с таблицей данных.
#     """
#     try:
#         # Загружаем файл из Bitrix24
#         file_name = FILE_NAME
#         file_data = get_word_file(file_name)
#
#         # Читаем Excel-файл с помощью pandas
#         df = pd.read_excel(file_data)
#
#         # Преобразуем DataFrame в HTML-таблицу
#         table_html = df.to_html(classes="table table-striped", border=0, index=False)
#
#         # Возвращаем HTML-страницу с таблицей
#         return templates.TemplateResponse(
#             "index.html",
#             {"request": {}, "file_name": file_name, "table_html": table_html}
#         )
#     except Exception as e:
#         # В случае ошибки возвращаем сообщение об ошибке
#         raise HTTPException(status_code=500, detail=f"Произошла ошибка: {str(e)}")


@app.post("/bitrix-webhook")
async def webhook(request: Request) -> JSONResponse:
    """
    Обрабатывает POST-запросы от Bitrix24 и отправляет ответ на основе данных из Excel.
    Также обрабатывает команду для приветственного сообщения и отправку кнопки.
    :param request: Входящий HTTP-запрос.
    :return: JSON-ответ с результатом обработки.
    """
    try:
        # Определяем тип контента
        content_type = request.headers.get("content-type")
        logger.info(f"Получен Content-Type: {content_type}")

        if "application/json" in content_type:
            data = await request.json()
            logger.info("JSON-данные запроса:")
            logger.info(data)
            # TODO: если нужно, реализуйте парсинг JSON здесь
            return JSONResponse(
                content={"status": "error", "message": "JSON пока не поддерживается"},
                status_code=400
            )

        elif "application/x-www-form-urlencoded" in content_type:
            form_data = await request.form()
            logger.info("Form-данные запроса:")
            logger.info(form_data)

            # Получаем параметры из формы
            message_text = form_data.get('data[PARAMS][MESSAGE]', '').strip('\r\n\t ').lower()
            user_id = form_data.get('data[PARAMS][FROM_USER_ID]', '')
            chat_id = form_data.get('data[PARAMS][CHAT_ID]', '') or form_data.get('data[PARAMS][TO_CHAT_ID]', '')
            bot_id = form_data.get('data[BOT][5732][BOT_ID]', '')
            auth_app_token = form_data.get('auth[application_token]', '')

            logger.info(f"message_text = '{message_text}'")
            logger.info(f"chat_id = '{chat_id}'")

            if not message_text or not chat_id:
                logger.warning(
                    f"Недостаточно данных для обработки запроса. MESSAGE='{message_text}', CHAT_ID='{chat_id}'")
                return JSONResponse(
                    content={"status": "error", "message": "Недостаточно данных"},
                    status_code=400
                )

            # Логика обработки приветственного сообщения
            greeting_keywords = {"привет", "начать", "hello", "hi"}
            if message_text in greeting_keywords:
                response_message = (
                    "Привет! 👋\n"
                    "Я помощник для поиска информации по материалам из базы Ф-ТД-008.\n"
                    "Чтобы начать, просто введите ID или заказной номер материала, и я предоставлю вам необходимую информацию.\n\n"
                    "По вопросам работы инструмента, улучшениям и предложениям обращайтесь к Гаврилову Михаилу.\n"
                    "📧E-mail: Gavrilov.Mikhail@vsmservice.ru"
                )
                send_message_to_bitrix(user_id, response_message, bot_id, auth_app_token)
                return JSONResponse(content={"status": "success", "message": "Приветственное сообщение отправлено"},
                                    status_code=200)

            # Если это первый запрос (например, пустое сообщение), отправляем кнопку
            if not message_text:
                send_button_to_bitrix(chat_id, bot_id, auth_app_token)
                return JSONResponse(content={"status": "success", "message": "Кнопка отправлена"}, status_code=200)

            # Логика обработки обычного запроса
            response_message = "Не найдено совпадений.\nЕсли вы видите это сообщение, значит запрашиваемый вами материал не внесен в базу Ф-ТД-008."
            for sheet_name, row in excel_data:
                # Проверяем полное совпадение с каждой ячейкой в строке
                for cell in row:
                    if str(cell).strip().lower() == message_text:
                        response_message = format_excel_row(sheet_name, row)
                        break
                if response_message != "Не найдено совпадений.":
                    break

            # Отправляем ответ в Bitrix24
            send_message_to_bitrix(user_id, response_message, bot_id, auth_app_token)
            return JSONResponse(content={"status": "success", "message": "Сообщение обработано"}, status_code=200)

        else:
            logger.warning(f"Неизвестный Content-Type: {content_type}")
            return JSONResponse(
                content={"status": "error", "message": "Неподдерживаемый формат данных"},
                status_code=415
            )

    except Exception as e:
        logger.error(f"Ошибка при обработке запроса: {e}")
        return JSONResponse(
            content={"status": "error", "message": "Внутренняя ошибка сервера"},
            status_code=500
        )


# Функция для отправки сообщений в чат Bitrix24
def send_message_to_bitrix(chat_id, message, bot_id, auth_app_token):
    """
        Отправляет сообщение в чат Bitrix24 через API.

        :param chat_id: ID чата или пользователя.
        :param message: Текст сообщения.
        :param bot_id: ID бота.
        :param auth_app_token: Токен авторизации приложения.
    """
    payload = {
        "BOT_ID": bot_id,
        "DIALOG_ID": str(chat_id),
        "MESSAGE": message,
        "CLIENT_ID": auth_app_token,
    }

    try:
        response = requests.post(f"{INCOMING_URL}/imbot.message.add", json=payload)
        if response.status_code != 200:
            logger.error(f"Ошибка при отправке сообщения в Bitrix24: {response.text}")
    except Exception as e:
        logger.error(f"Ошибка при отправке сообщения в Bitrix24: {e}")


def send_button_to_bitrix(chat_id, bot_id, auth_app_token):
    """
    Отправляет кнопку "Привет" в чат Bitrix24 через API.
    :param chat_id: ID чата или пользователя.
    :param bot_id: ID бота.
    :param auth_app_token: Токен авторизации приложения.
    """
    payload = {
        "BOT_ID": bot_id,
        "DIALOG_ID": str(chat_id),
        "MESSAGE": "Нажмите кнопку ниже, чтобы начать:",
        "KEYBOARD": {
            "BUTTONS": [
                {
                    "TEXT": "Привет",
                    "COMMAND": "привет",  # Команда, которая будет отправлена при нажатии
                    "BG_COLOR": "#2961c2",  # Цвет фона кнопки
                    "TEXT_COLOR": "#fff"  # Цвет текста
                }
            ]
        },
        "CLIENT_ID": auth_app_token,
    }
    try:
        response = requests.post(f"{INCOMING_URL}/imbot.message.add", json=payload)
        if response.status_code != 200:
            logger.error(f"Ошибка при отправке кнопки в Bitrix24: {response.text}")
    except Exception as e:
        logger.error(f"Ошибка при отправке кнопки в Bitrix24: {e}")


# Запуск сервера
if __name__ == "__main__":
    uvicorn.run(app, host="localhost", port=5000)
