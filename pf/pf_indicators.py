#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import postgresql

project_id = 5

# Начало проекта
start_year = 2011
# Количество периодов проекта (на сколько лет проект)
period_count = 20

db = postgresql.open('pq://projectfinance:projectfinance@91.241.45.10:5432/postgres')
rb = xlrd.open_workbook('h:/indicators.xlsx')
sheet = rb.sheet_by_index(0)


# name: это строка которую транслитим
def translit(name):
    """
    Автор: LarsKort
    Дата: 16/07/2011; 1:05 GMT-4;
    Не претендую на "хорошесть" словарика. В моем случае и такой пойдет,
    вы всегда сможете добавить свои символы и даже слова. Только
    это нужно делать в обоих списках, иначе будет ошибка.
    """
    # Слоаврь с заменами
    slovar = {'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e',
              'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'i', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n',
              'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'h',
              'ц': 'c', 'ч': 'cz', 'ш': 'sh', 'щ': 'scz', 'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e',
              'ю': 'u', 'я': 'ja', 'А': 'a', 'Б': 'b', 'В': 'v', 'Г': 'g', 'Д': 'd', 'Е': 'e', 'Ё': 'e',
              'Ж': 'zh', 'З': 'z', 'И': 'i', 'Й': 'i', 'К': 'k', 'Л': 'l', 'М': 'm', 'Н': 'n',
              'О': 'o', 'П': 'p', 'Р': 'r', 'С': 's', 'Т': 't', 'У': 'u', 'Ф': 'F', 'х': 'h',
              'Ц': 'c', 'Ч': 'cz', 'Ш': 'sh', 'Щ': 'scz', 'Ъ': '', 'Ы': 'y', 'Ь': '', 'Э': 'e',
              'Ю': 'u', 'Я': 'ja', ',': '', '?': '', ' ': '_', '~': '', '!': '', '@': '', '#': '',
              '$': '', '%': '', '^': '', '&': '', '*': '', '(': '', ')': '', '-': '', '=': '', '+': '',
              ':': '', ';': '', '<': '', '>': '', '\'': '', '"': '', '\\': '', '/': '', '№': '',
              '[': '', ']': '', '{': '', '}': '', 'ґ': '', 'ї': '', 'є': '', 'Ґ': 'g', 'Ї': 'i',
              'Є': 'e'}

    # Циклически заменяем все буквы в строке
    for key in slovar:
        name = name.replace(key, slovar[key])
    return name

# функция создания индикатора (в t_indicatior)
def create_indicator(id, name, unit_of_measure):
    code = translit(name)
    str = "INSERT INTO t_indicator(id, code, name, description, indicator_type, period_type, unit_of_measure, project_id) VALUES ({0}, '{1}', '{2}', '{2}', 'STATIC', 'YEAR', '{3}', {4});".format(
        id, code, name, unit_of_measure, project_id)
    print("\n" + str)

def create_indicator_value(id_value, id, plan_value, fact_value, estimate_value, period):
    if (fact_value != ''):
        str = "INSERT INTO t_indicator_value(id, indicator_id, plan_value, actual_value, estimated_value, time_period_id) VALUES ({0}, {1}, '{2}', '{3}', '{4}', {5});".format(
            id_value, id, plan_value, fact_value, estimate_value, period)
    else:
        str = "INSERT INTO t_indicator_value(id, indicator_id, plan_value, estimated_value, time_period_id) VALUES ({0}, {1}, '{2}', '{3}', '{4}');".format(
            id_value, id, plan_value, estimate_value, period)
    print(str)
    #print("{0}".format(3*(i-1)+4))

def create_indicator_catalog_item(id_catalog_item, id, parent_id, title, project_id):
    item_index_sql = db.query("SELECT MAX(item_index) FROM t_indicator_catalog_item WHERE parent_id = {0} AND project_id = {1};".format(parent_id, project_id));
    try:
        item_index = int(item_index_sql[0][0]) + 1
    except Exception:
        item_index = 1
    code = translit(title)
    str = "INSERT INTO t_indicator_catalog_item(id, indicator_id, parent_id, title, code, project_id, item_index) VALUES ({0}, {1}, {2}, '{3}', '{4}', {5}, {6});".format(
        id_catalog_item, id, parent_id, title, code, project_id, item_index)
    print(str)

id_sql = db.query("SELECT MAX(id) FROM t_indicator;");
id = int (id_sql[0][0]) + 1

id_catalog_item_sql = db.query("SELECT MAX(id) FROM t_indicator_catalog_item;");
id_catalog_item = int (id_catalog_item_sql[0][0]) + 1

id_value_sql = db.query("SELECT MAX(id) FROM t_indicator_value;");
id_value = int (id_value_sql[0][0]) + 1

# Для каждой строки файла
for rownum in range(sheet.nrows):
    row = sheet.row_values(rownum)
    if (row[0] != 0):
        # создаем индикатор
        create_indicator(id, row[0], row[1])
        create_indicator_catalog_item(id_catalog_item, id, int(row[2]), row[0], project_id)
        id_catalog_item += 1
        # Создаем indicator_value для каждого периода
        for i in range(0, 20):
           year_sql = db.query("SELECT id FROM t_time_period WHERE period_value = {0};".format(start_year+i));
           year = int (year_sql[0][0])
           create_indicator_value(id_value, id, row[(3*i + 3)], row[(3*i + 4)], row[(3*i + 5)], year)
           id_value += 1
        id += 1




