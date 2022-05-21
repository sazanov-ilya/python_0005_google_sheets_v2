import os.path
from googleapiclient.discovery import build
from google.oauth2 import service_account
import datetime
# from datetime import date

# Импорт модуля класса БД
from sql import Sql
# Импорт модуля для расчета диапазона
from a1range import A1Range

# Получаем логины и пароли для google sheets
from settings import GOOGLE_SAMPLE_SPREADSHEET_ID


# Полные права на сервисный аккаунт google
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# Абсолютный путь к проекту
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Путь до файла настроек подключения в формате json
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'credentials.json')
# Устанавливаем параметры для подключения к google sheet
CREDENTIALS = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
# Настройка подключения к google sheet
SERVICE = build('sheets', 'v4', credentials=CREDENTIALS).spreadsheets().values()

# Идентификатор электронной таблицы (из ссылки)
SAMPLE_SPREADSHEET_ID = GOOGLE_SAMPLE_SPREADSHEET_ID
# Диапазон данных на листе
# SAMPLE_RANGE_NAME = 'Class Data!A2:E'
SHEET_LIST_NAME = 'Лист1'  # Название листа (данные со всей страницы)


def main():
    # ######################### #
    # Чтение данных со страницы #
    # ######################### #

    # Получаем словарь с данными с листа google sheet
    result = SERVICE.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                         range=SHEET_LIST_NAME).execute()

    # # Перебираем все ключи словаря
    # for key, values in result.items():
    #     print(f'{key}: {values}')

    # Содержанием таблицы
    table_from_sheet = result.get('values', [])

    # # Удаляем пустые строки (если нужно, но тогда потребуется полное обновление всего отчета)
    # for index, row in enumerate(table_from_sheet):
    #     # print(f'{index}: {value}')
    #     if len(row) == 0:
    #         table_from_sheet.pop(index)
    #         # print(f'{index}: {value}')

    # ###################### #
    # Получение данные из БД #
    # ###################### #
    sql = Sql('oktell')  # Создаем экземпляр класса БД
    table_report = sql.get_report()  # Получаем данные отчета

    # ############# #
    # Запись данных #
    # ############# #
    VALUE_INPUT_OPTION = 'USER_ENTERED'  # Тип записи данных на лист google sheet

    # Если еще нет, создаем заголовок
    if len(table_from_sheet) == 0:
        table_header_to_sheet = {'values': [
            ['Оператор',
             'Задача',
             'Время в системе',
             'Время разговора',
             'Ср. время разговора, с.',
             'Время постобработки звонка',
             'Кофе-брейк',
             'Обед',
             'Технический',
             'Обратная связь',
             'Вертушка',
             'Ивент',
             'Оффлайн задачи',
             'Кол-во звонков',
             'Кол-во контактов',
             'Кол-во звонков к контакту',
             'Дозвон',
             '% дозвона',
             'Дата']]
        }
        # Диапазон для заголовка
        # RANGE = 'Лист1!A8:B9'
        range_header = A1Range.create_a1range_from_list(SHEET_LIST_NAME, 1, 1, table_header_to_sheet['values']).format()
        # Пишем заголовок
        request = SERVICE.update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                 range=range_header,
                                 valueInputOption=VALUE_INPUT_OPTION,
                                 body=table_header_to_sheet)
        response_header = request.execute()

    # (собираем нужный формат словаря)
    # ARRAY = {'values': [[0, 8], [1, 2], ['qwe', 'rew']]}
    table_to_sheet = {}
    table_to_sheet['values'] = table_report
    table_to_sheet['values'] = list(map(list, table_report))  # преобразуем список кортежей к списку словарей
    # print(table_to_sheet)

    # Поучаем диапазон для выгрузки/обновления данных
    # RANGE = 'Лист1!A8:B9'
    # Диапазон по умолчанию со второй строки
    range_data = A1Range.create_a1range_from_list(SHEET_LIST_NAME, 2, 1, table_to_sheet['values']).format()
    # Корректируем диапазон, если уже есть данные
    if len(table_from_sheet) > 1:  # Значит должен быть заголовок и какие-то данные
        print('Переопределяем диапазон')

        current_row_index = len(table_from_sheet)  # По умолчанию добавляем в конец таблицы
        current_date = datetime.date.today()  # Текущая дата

        # Проверяем и обновляем данные за сегодня
        for row_index in range(1, len(table_from_sheet)):
            # row_value = table_from_sheet[row_index]  # Содержание строки
            # print(row_index, row_value)
            # for coll_index, col_value in enumerate(row_value):
            #     print(f'{coll_index}: {col_value}')
            # print(table_from_sheet[row_index][18])

            # Игнорируем строки без даты
            try:
                row_date = datetime.datetime.strptime(table_from_sheet[row_index][18], '%Y-%m-%d').date()  # Дата
            except:
                continue
            if row_date == current_date:
                current_row_index = row_index  # Переопределяем индекс
                break
        # print(f'current_row_index: {current_row_index}')
        range_data = A1Range.create_a1range_from_list(SHEET_LIST_NAME, current_row_index + 1, 1, table_to_sheet['values']).format()

    # Пишем содержание таблицы
    request = SERVICE.update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                             range=range_data,
                             valueInputOption=VALUE_INPUT_OPTION,
                             body=table_to_sheet)
    response_table = request.execute()


if __name__ == '__main__':
    main()