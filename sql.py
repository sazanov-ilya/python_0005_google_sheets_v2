import pyodbc
from datetime import datetime

# Получаем логины и пароли для БД
from settings import SQL_SERVER_NAME, SQL_LOGIN, SQL_PASSWORD


class Sql:
    """ Класс подключения к БД """
    def __init__(self, database, server=SQL_SERVER_NAME):
        # Подключаемся а БД
        self.conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                                   'SERVER=' + server + ';'
                                   'DATABASE=' + database + ';'
                                   'UID=' + SQL_LOGIN + ';'
                                   'PWD=' + SQL_PASSWORD + ';'
                                   'Trusted_Connection=no;')
        # Создаем курсор для чтения
        self.cursor = self.conn.cursor()

        self.query = "-- {}\n\n-- Made in Python".format(datetime.now().strftime("%d/%m/%Y"))

    def insert_data(self, id, data):
        """ Процедура сохранения данных в БД """
        try:
            self.conn.execute(
                '''INSERT INTO tbl_test_python(id, data) VALUES(?, ?)''',
                (id, data)
            )
            self.conn.commit()
        except pyodbc.Error as err:
            print(err.__str__())

    def get_data(self):
        """ Процедура получения данных из БД через select """
        try:
            self.cursor.execute(
                '''SELECT [id],[data] FROM _test_python'''
            )
        except pyodbc.Error as err:
            print(err.__str__())

        tables = self.cursor.fetchall()
        return tables

    def get_report(self):
        """ Процедура получения данных из БД через exec """
        try:
            self.cursor.execute(
                '''exec oktell.dbo.Report_on_Operators_V3'''
            )
        except pyodbc.Error as err:
            print(err.__str__())

        # Список всех строк
        table = self.cursor.fetchall()
        # Одна строка
        # table = self.cursor.fetchone()
        return table

    # def push_dataframe(self, data, table="raw_data", batchsize=500):
    #     """ Функция push_dataframe позволит поместить в базу данных датафрейм Pandas.
    #      https://proglib.io/p/pandas-tricks """
    #     cursor = self.cnxn.cursor()  # создаем курсор
    #     cursor.fast_executemany = True  # активируем быстрое выполнение
    #
    #     # создаём заготовку для создания таблицы (начало)
    #     query = "CREATE TABLE [" + table + "] (\n"
    #
    #     # итерируемся по столбцам
    #     for i in range(len(list(data))):
    #         query += "\t[{}] varchar(255)".format(list(data)[i])  # add column (everything is varchar for now)
    #         # добавляем корректное завершение
    #         if i != len(list(data)) - 1:
    #             query += ",\n"
    #         else:
    #             query += "\n);"
    #
    #     cursor.execute(query)  # запуск создания таблицы
    #     self.cnxn.commit()  # коммит для изменений
    #
    #     # append query to our SQL code logger
    #     self.query += ("\n\n-- create table\n" + query)
    #
    #     # вставляем данные в батчи
    #     query = ("INSERT INTO [{}] ({})\n".format(table,
    #                                               '[' + '], ['  # берем столбцы
    #                                               .join(list(data)) + ']') +
    #              "VALUES\n(?{})".format(", ?" * (len(list(data)) - 1)))
    #
    #     # вставляем данные в целевую таблицу
    #     for i in range(0, len(data), batchsize):
    #         if i + batchsize > len(data):
    #             batch = data[i: len(data)].values.tolist()
    #         else:
    #             batch = data[i: i + batchsize].values.tolist()
    #         # запускаем вставку батча
    #         cursor.executemany(query, batch)
    #         self.cnxn.commit()