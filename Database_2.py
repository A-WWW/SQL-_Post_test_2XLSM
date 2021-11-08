import openpyxl
from psycopg2 import OperationalError
from psycopg2 import Error
import psycopg2
from Config import host, user, password, db_name
try:
    connection = psycopg2.connect(
        host=host,
        user=user,
        password=password,
        database=db_name
    )
    connection.autocommit = True
    with connection.cursor() as cursor:
        cursor.execute(
            "SELECT version();"
        )
        print(f"Server version: {cursor.fetchone()}")
        print(connection.get_dsn_parameters(), "\n")
# таблицы если верить справочным мтериалам приведены до втотого уровня нормальности, рекомендуется доводить как мимнимум
# до третьего (в данном случае это сделано частично), полностью делать это думаю будет излишне так как при большом количестве материалов будет сильно тормозить  базу.

    # with connection.cursor() as cursor:
    #     cursor.execute(
    #         '''CREATE TABLE country
    #                (Id SERIAL PRIMARY KEY,
    #                Country CHARACTER VARYING(50) UNIQUE NOT NULL
    #                  );'''
    #     )
    #     print("(INFO) Table created country")


    # with connection.cursor() as cursor:
    #     cursor.execute(
    #         '''CREATE TABLE region
    #             (Id SERIAL PRIMARY KEY,
    #             region CHARACTER VARYING(50) UNIQUE NOT NULL,
    #             countryId INTEGER REFERENCES country (Id) ON DELETE RESTRICT
    #                             );'''
    #     )
    #     print("(INFO) Table created region")

# many-to-many, поэтому есть промежуточная таблица city_1
#     with connection.cursor() as cursor:
#         cursor.execute(
#             '''CREATE TABLE city
#                 (Id SERIAL PRIMARY KEY,
#                 city CHARACTER VARYING(50) UNIQUE NOT NULL
#                                                );'''
#         )
#         print("(INFO) Table created city")


    # with connection.cursor() as cursor:
    #     cursor.execute(
    #         '''CREATE TABLE city_1
    #              (Id SERIAL PRIMARY KEY,
    #              cityId INTEGER REFERENCES city(Id) ON DELETE RESTRICT,
    #              regionId INTEGER REFERENCES region (Id) ON DELETE RESTRICT
    #                                               );'''
    #     )
    #     print("(INFO) Table created city_1")


    # with connection.cursor() as cursor:
    #     cursor.execute(
    #         '''CREATE TABLE language
    #             (Id SERIAL PRIMARY KEY,
    #             language CHARACTER VARYING(30) UNIQUE NOT NULL
    #                                             );'''
    #     )
    #     print("(INFO) Table created language")

    # with connection.cursor() as cursor:
    #     cursor.execute(
    #         '''CREATE TABLE mass_media
    #         (Id SERIAL PRIMARY KEY,
    #         mass_media CHARACTER VARYING(30) UNIQUE NOT NULL
    #                                         );'''
    #     )
    #     print("(INFO) Table created mass_media")

    # в дате не формат даты и времени так как все настойчиво рекомендуют использовать строку так как поиск организовать в будущем будет проще.
    # with connection.cursor() as cursor:
    #     cursor.execute(
    #         '''CREATE TABLE news_data
    #         (Id SERIAL PRIMARY KEY,
    #         heading text,
    #         text text,
    #         mass_mediaId INTEGER REFERENCES mass_media(Id) ON DELETE RESTRICT,
    #         source CHARACTER VARYING(80),
    #         city_1Id INTEGER REFERENCES city_1(Id)ON DELETE CASCADE,
    #         pub_level  CHARACTER VARYING(50),
    #         title CHARACTER VARYING(50),
    #         languageId INTEGER REFERENCES language(Id) ON DELETE RESTRICT,
    #         rel_date CHARACTER VARYING(50),
    #         link text,
    #         circulation integer
    #         );'''
    #     )
    #     print("(INFO) Table created news_data")




# парсер файла как пример, вытащено все но таблицы будут заполнятся по выделенным столбцам
# book = openpyxl.open("test_news.xlsm", read_only=True)
# wb = book.active
# print(wb["B2"].value)
# for i in range(1, wb.max_row + 1):
#     heading = wb[i][0].value
#     text = wb[i][1].value
#     mass_media = wb[i][2].value
#     source = wb[i][3].value
#     country = wb[i][4].value
#     region = wb[i][5].value
#     city = wb[i][6].value
#     pub_level = wb[i][7].value
#     title = wb[i][8].value
#     language = wb[i][9].value
#     rel_date = wb[i][11].value
#     link = wb[i][12].value
#     circulation = wb[i][13].value
#     print(i, heading, text, mass_media, source, country, region, city, pub_level, title, language, rel_date, link,
#           circulation)
# for i in wb.iter_rows(min_row=2, max_row=5, min_col=4, max_col=7):
#     for cell in i:
#         print(i.value, end=" ")
#     print()
# sheets = book.worksheets
# print(sheets)
# sheets2 = book.worksheets[2]
# print(sheets2['A2'])

except Exception as _ex:
    print("(INFO) Error wile", _ex)
finally:
    if connection:
        #cursor.close()
        connection.close()
        print("[INFO] conection clossed")
