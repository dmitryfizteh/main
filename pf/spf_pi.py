# Скрипт наполнения данных Project indicators для СПФ

#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import postgresql
import json
import sys
from datetime import datetime, date, time

undo = ""
sql_update=""
sql_insert =""

project_id = 1
#sql="INSERT INTO spf.project(id, name) VALUES (31,'Темповый проект');"

pi_item_id = 1000
pi_catalog_id = 1
cat_item_id = 5000
pi_version_id = 1000

# TODO: Создать периоды проекта
# Перевод из периодов файла загрузки в периоды проекта в БД
periods = {1:8,2:9,3:11,4:12,5:13,6:14,7:15,8:16,9:17,10:18,11:19,12:20,13:21,14:22,15:23,16:24,17:25,18:26,19:27,20:28}

'''
def run_sql(sql, sql_undo):
    global undo, sql_insert
    sql_insert += sql + "\n"
    undo = sql_undo + undo

# Создаем PI и фактические значение
def create_pi_item(id_pi_item, name, code, layeredattrs, projectid):
    sql = "INSERT INTO spf.projectindicator(id, name, code, layeredattrs, project_id) VALUES ({0}, '{1}', '{2}', '{3}', '{4}');".format(
        id_pi_item, name, code, layeredattrs, projectid)
    sql_undo = "DELETE FROM spf.projectindicator WHERE id={0};\n".format(id_pi_item)
    run_sql(sql, sql_undo)
    return id_pi_item+1

# Создаем "версию" плана PI
def create_piversion(id, master_id, layeredattrs, versionid=1):
    sql = "INSERT INTO spf.projectindicatorversion(id, layeredattrs, versionid, master_id) VALUES ({0}, '{1}', {2}, {3});".format(
        id, layeredattrs, versionid, master_id)
    sql_undo = "DELETE FROM spf.projectindicatorversion WHERE id={0};\n".format(id)
    return id+1

# Создаем запись в каталоге PI
def create_cat_item(id, parent_id, item_id, name, code, catalog_id=pi_catalog_id, versionid=1):
    if (item_id):
        sql = "INSERT INTO spf.projectindicatorcatalogitem(id, versionid, catalog_id, parent_id, item_id) VALUES ({0}, {1}, {2}, {3}, {4});".format(
            id, versionid, catalog_id, parent_id, item_id)
    else:
        if (parent_id):
            sql = "INSERT INTO spf.projectindicatorcatalogitem(id, code, name, versionid, catalog_id, parent_id) VALUES ({0}, '{1}', '{2}', {3}, {4}, {5});".format(
                id, code, name, versionid, catalog_id, parent_id)
        else:
            sql = "INSERT INTO spf.projectindicatorcatalogitem(id, code, name, versionid, catalog_id) VALUES ({0}, '{1}', '{2}', {3}, {4});".format(
                id, code, name, versionid, catalog_id)
    sql_undo = "DELETE FROM spf.projectindicatorcatalogitem WHERE id={0};\n".format(id) + undo
    run_sql(sql, sql_undo)
    return id+1

'''
# Создаем PI и фактические значение
def create_pi_item(id_pi_item, name, code, layeredattrs, projectid):
    global undo, sql_insert
    sql = "INSERT INTO spf.projectindicator(id, name, code, layeredattrs, projectid) VALUES ({0}, '{1}', '{2}', '{3}', '{4}');".format(
        id_pi_item, name, code, layeredattrs, projectid)
    sql_insert += sql + "\n"
    # print(sql)

    undo = "DELETE FROM spf.projectindicator WHERE id={0};\n".format(id_pi_item) + undo

    id_pi_item += 1
    return id_pi_item

# Создаем "версию" плана PI
def create_piversion(id, master_id, layeredattrs, versionid=1):
    global undo, sql_insert
    sql = "INSERT INTO spf.projectindicatorversion(id, layeredattrs, versionid, master_id) VALUES ({0}, '{1}', {2}, {3});".format(
        id, layeredattrs, versionid, master_id)
    sql_insert += sql + "\n"
    # print(sql)

    undo = "DELETE FROM spf.projectindicatorversion WHERE id={0};\n".format(id) + undo

    id += 1
    return id

# Создаем запись в каталоге PI
def create_cat_item(id, parent_id, item_id, name, code, catalog_id=pi_catalog_id, versionid=1):
    global undo, sql_insert
    if (item_id):
        sql = "INSERT INTO spf.projectindicatorcatalogitem(id, versionid, catalog_id, parent_id, item_id) VALUES ({0}, {1}, {2}, {3}, {4});".format(
            id, versionid, catalog_id, parent_id, item_id)
    else:
        if (parent_id):
            sql = "INSERT INTO spf.projectindicatorcatalogitem(id, code, name, versionid, catalog_id, parent_id) VALUES ({0}, '{1}', '{2}', {3}, {4}, {5});".format(
                id, code, name, versionid, catalog_id, parent_id)
        else:
            sql = "INSERT INTO spf.projectindicatorcatalogitem(id, code, name, versionid, catalog_id) VALUES ({0}, '{1}', '{2}', {3}, {4});".format(
                id, code, name, versionid, catalog_id)
    sql_insert += sql + "\n"
    # print(sql)

    undo = "DELETE FROM spf.projectindicatorcatalogitem WHERE id={0};\n".format(id) + undo

    id += 1
    return id

def update_pi_item(id, layeredattrs, sql_update):
    sql_update = sql_update + "UPDATE spf.projectindicator SET layeredattrs='{1}' WHERE id={0};".format(id, layeredattrs)
    return sql_update

rb = xlrd.open_workbook('data/indicators.xlsx')
sheet = rb.sheet_by_name("PI")

# TODO: сделать обработку всего файла
for rownum in range(2, sheet.nrows):
    row = sheet.row_values(rownum)

    if (row[0] != ""):
        num = str(row[0]).split(";")

        if (len(num) == 1):
            cat_item_id = create_cat_item(cat_item_id, "", "", row[1], "pici_" + str(cat_item_id))
            parent_tag1 = cat_item_id - 1
        if(len(num) == 2):
            cat_item_id = create_cat_item(cat_item_id, parent_tag1, "", row[1], "pici_" + str(cat_item_id))
            parent_tag2 = cat_item_id - 1
        if (len(num) == 3):
            cat_item_id = create_cat_item(cat_item_id, parent_tag2, "", row[1], "pici_" + str(cat_item_id))
            parent_tag3 = cat_item_id - 1
        if (len(num) == 4):
            cat_item_id = create_cat_item(cat_item_id, parent_tag3, "", row[1], "pici_" + str(cat_item_id))
        last_tag = cat_item_id - 1

    if (row[0] == ""):
        fact = {}
        update_fact = {}
        plan = {}
        for i in range(1, 20):
            key = periods[i]
            # До какого периода вносить фактические данные
            if (i < 10):
                fact[key] = row[2 * (i - 1) + 4]
            # До какого периода вносить фактические данные (с помощью UPDATE)
            if (i < 11):
                update_fact[key] = row[2 * (i - 1) + 4]
            plan[key] = row[2 * (i - 1) + 3]
            i += 1
        layeredattrs = json.dumps({"FACT": fact})
        layeredattrs_update_fact = json.dumps({"FACT": update_fact})
        layeredattrs_plan = json.dumps({"PLAN": plan})
        code = "pi_"+str(pi_item_id)
        pi_item_id = create_pi_item(pi_item_id, row[1], code, layeredattrs, project_id)
        sql_update = update_pi_item(pi_item_id - 1, layeredattrs_update_fact, sql_update)
        pi_version_id = create_piversion(pi_version_id, pi_item_id - 1, layeredattrs_plan)
        cat_item_id = create_cat_item(cat_item_id, last_tag, pi_version_id - 1, "", "")

#print (undo)
#print (sql_update)

undo_file = open('results/insert_pi_'+ datetime.now().strftime('%Y%m%d_%H%M') +'.txt', 'w')
undo_file.write(sql_insert)
undo_file.close()

undo_file = open('results/undo_pi_'+ datetime.now().strftime('%Y%m%d_%H%M') +'.txt', 'w')
undo_file.write(undo)
undo_file.close()

undo_file = open('results/update_pi_'+ datetime.now().strftime('%Y%m%d_%H%M') +'.txt', 'w')
undo_file.write(sql_update)
undo_file.close()

# В цвете вывод информации о завершении работы
OKGREEN = '\033[94m'
ENDC = '\033[0m'
print(OKGREEN  + "Скрипт {} закончил работу".format(sys.argv[0])+ ENDC)