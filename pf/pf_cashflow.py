#!/usr/bin/env python
# -*- coding: utf-8 -*-

from datetime import datetime, date, time
import xlrd
import postgresql

project_id = 5

# Начало проекта
start_year = 2011
# Количество периодов проекта (на сколько лет проект)
period_count = 20

undo = ""
asset_value =""

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

def create_cf(id_cf, title, cf_direction, project_stage):
    str = "INSERT INTO t_cash_flow(id, title, project_id, cf_source_code, cf_direction, cf_type, project_stage) VALUES ({0}, '{1}', {2}, '{3}', '{4}'," \
          " '{5}', '{6}');".format(id_cf, title, project_id, title, cf_direction, "PROJECT_STAGE", project_stage)
    print(str)
    global undo
    undo = "DELETE FROM t_cash_flow WHERE id={0};\n".format(id_cf) + undo
    id_cf += 1
    return id_cf

# Добавляем CF_item (id_item, title, item_sum, time_preiod_id, cash_flow_id)
def create_cf_item(id_cf_item, title_cf_item, item_sum, time_period_id, cash_flow_id, last_tag):
    if(item_sum ==""):
        item_sum = 0
    str = "INSERT INTO t_cash_flow_item(id, title, item_sum, time_period_id, cash_flow_id) VALUES ({0}, '{1}', {2}, {3}, {4});".format(
        id_cf_item, title_cf_item, item_sum, time_period_id, cash_flow_id)
    print(str)
    global undo
    undo = "DELETE FROM t_cash_flow_item WHERE id={0};\n".format(id_cf_item) + undo

    str = "INSERT INTO t_cash_flow_tag_binding(cf_tag_code, cf_item_id) VALUES ('{0}', {1});".format(last_tag,id_cf_item)
    print(str)
    undo = "DELETE FROM t_cash_flow_tag_binding WHERE cf_tag_code='{0}' AND cf_item_id={1};\n".format(last_tag,id_cf_item) + undo
    id_cf_item += 1
    return id_cf_item

def create_cf_items(id_cf_item, id_cf, row, excel_arr, last_tag):
    for i in range(0, period_count):
        year_sql = db.query("SELECT id FROM t_time_period WHERE period_value = {0};".format(start_year + i));
        year = int(year_sql[0][0])
        id_cf_item = create_cf_item(id_cf_item, row[1], row[(3 * i + 4 + excel_arr)], year, id_cf, last_tag)
    return id_cf_item

f = open('passwd.txt')
for connect_str in f:
    db = postgresql.open(connect_str)
f.close()

id_cf_sql = db.query("SELECT MAX(id) FROM t_cash_flow;");
id_cf = int (id_cf_sql[0][0]) + 1

id_cf_item_sql = db.query("SELECT MAX(id) FROM t_cash_flow_item;");
id_cf_item = int (id_cf_item_sql[0][0]) + 1

id_asset_part_sql = db.query("SELECT MAX(id) FROM t_asset_part;");
id_asset_part = int (id_asset_part_sql[0][0]) + 1

id_asset_value_sql = db.query("SELECT MAX(id) FROM t_asset_value;");
id_asset_value = int (id_asset_value_sql[0][0]) + 1

id_asset_income_sql = db.query("SELECT MAX(id) FROM t_asset_income;");
id_asset_income = int (id_asset_income_sql[0][0]) + 1

id_asset_value_item_sql = db.query("SELECT MAX(id) FROM t_asset_value_item;");
id_asset_value_item = int (id_asset_value_item_sql[0][0]) + 1

last_row = ""
rb = xlrd.open_workbook('h:/indicators.xlsx')
sheet = rb.sheet_by_index(1)

def create_tag(tag_name, parent_tag):
    if(parent_tag != ""):
        str = "INSERT INTO t_cash_flow_tag(code, description, parent_code) VALUES ('{0}', '{1}', '{2}');".format(
            translit(tag_name), tag_name, parent_tag)
    else:
        str = "INSERT INTO t_cash_flow_tag(code, description) VALUES ('{0}', '{1}');".format(
            translit(tag_name), tag_name)
    print(str)
    global undo
    undo = "DELETE FROM t_cash_flow_tag WHERE code='{0}';\n".format(translit(tag_name)) + undo

for rownum in range(2, sheet.nrows):
    row = sheet.row_values(rownum)

    if (row[0] != "cf" and row[0] != "av" and row[0] != "ap" and row[0] != "cv" and row[0] != "cp" and row[0] != ""):
        num = row[0].split(";")
        if (len(num) == 1):
            create_tag(row[1], "")
            parent_tag1 = translit(row[1])
        if(len(num) == 2):
            create_tag(row[1], parent_tag1)
            parent_tag2 = translit(row[1])
        if (len(num) == 3):
            create_tag(row[1], parent_tag2)
            parent_tag3 = translit(row[1])
        if (len(num) == 4):
            create_tag(row[1], parent_tag3)

        last_tag = translit(row[1])

    # Создаем CF (cash_flow_id, title, project_id, cf_source_code = title, cf_direction, cf_type = PROJECT_STAGE, project_stage
    # = PLAN/FACT/ESTIMATE)
    if (row[0] == "cf"):
        id_cf = create_cf(id_cf, row[1], row[2], "PLAN")
        id_cf_item = create_cf_items(id_cf_item, id_cf-1, row, 0, last_tag) #Смещение в таблице для PLAN
        id_cf = create_cf(id_cf, row[1], row[2], "FACT")
        id_cf_item = create_cf_items(id_cf_item, id_cf-1, row, 1, last_tag)
        id_cf = create_cf(id_cf, row[1], row[2], "ESTIMATE")
        id_cf_item = create_cf_items(id_cf_item, id_cf-1, row, 2, last_tag)
        last_cf_item_id = id_cf_item-1

    if (row[0] == "ap"):
        for i in range(0, period_count): # х3
            # Читаем из файла значения цен и копируем в список
            asset_value = row

    if (row[0] == "av"):
        print("*" + row[1])
        t_asset_id = 2
        # создаем t_asset_part
        str = "INSERT INTO t_asset_part(id, title, launch_time_period_id, asset_id) VALUES ({0}, '{1}', {2}, {3});".format(
            id_asset_part, row[1], 16, t_asset_id)
        print(str)
        #global undo
        undo = "DELETE FROM t_asset_part_item WHERE id={0};\n".format(id_asset_part) + undo
        id_asset_part += 1

        for i in range(0, period_count): # х3
            # создаем из списка asset_value строку создания цен
            str = "INSERT INTO t_asset_value(id, title, asset_value) VALUES ({0}, '{1}', {2});".format(
                id_asset_value, asset_value[1], asset_value[3*i+4]) #План
            print(str)
            undo = "DELETE FROM t_asset_value WHERE id={0};\n".format(id_asset_value) + undo
            id_asset_value += 1

            # Создаем строку с объемом
            str = "INSERT INTO t_asset_income(id, title, asset_value, volume, asset_part_id, cash_flow_item_id) VALUES ({0}, '{1}', {2}, '{3}', {4}, {5});".format(
                id_asset_income, row[1], row[3 * i + 4], row[3], t_asset_id, last_cf_item_id)  # План TODO: проверить, возможно косяк с cf_item
            print(str)
            undo = "DELETE FROM t_asset_income WHERE id={0};\n".format(id_asset_income) + undo
            id_asset_income += 1

            year_sql = db.query("SELECT id FROM t_time_period WHERE period_value = {0};".format(start_year + i));
            year = int(year_sql[0][0])

            # Связываем цену и объем
            str = "INSERT INTO t_asset_value_item(id, asset_value_id, asset_income_id, conditions, asset_value_type, project_scenario_id, project_stage, time_period_id) " \
                  "VALUES ({0}, '{1}', {2}, '{3}', {4}, {5}, {6}, {7});".format(id_asset_value_item, id_asset_value-1, id_asset_income-1, "NO", "PROJECT_STAGE", 1, "PLAN", year)  # План
            print(str)
            undo = "DELETE FROM t_asset_value_item WHERE id={0};\n".format(id_asset_value_item) + undo
            id_asset_value_item += 1

undo_file = open('undo_cf_'+ datetime.now().strftime('%Y%m%d_%H%M%S') +'.txt', 'w')
undo_file.write(undo)
undo_file.close()

# Проверяем, что CF отсутствует (иначе полчаем cash_flow_id)
# Создаем CF (cash_flow_id, title, project_id, cf_source_code = title, cf_direction, cf_type = PROJECT_STAGE, project_stage
# = PLAN/FACT/ESTIMATE)

# Для каждого периода и CF (там плюс/минус и пл/ф/пр)
# Добавляем CF_item (id_item, title, item_sum, time_preiod_id, cash_flow_id)

# Создаем tag

# Связываем cf_tag_code и cf_item_id
