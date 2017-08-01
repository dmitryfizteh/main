# Скрипт наполнения данных CashFlow для СПФ

#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import postgresql
import json
import sys
from datetime import datetime, date, time

undo = ""
sql_insert =""
sql_update=""

project_id = 1
#sql="INSERT INTO spf.project(id, name) VALUES (31,'Темповый проект');"
currency_id = 1

# TODO: Запросить cf_item_id (max_id)
cf_item_id = 5601
cf_version_id = 5601
cf_cat_item_id = 5601

# TODO: Создать cashflowcatalog (max_id)
cashflowcatalog_id = 1

# TODO: Создать периоды проекта
# Перевод из периодов файла загрузки в периоды проекта в БД
periods = {1:8,2:9,3:11,4:12,5:13,6:14,7:15,8:16,9:17,10:18,11:19,12:20,13:21,14:22,15:23,16:24,17:25,18:26,19:27,20:28}

# Создаем CF и фактические значение
def create_cf_item(id_cf_item, name, code, layeredattrs, direction, type, currency_id):
    global undo, sql_insert
    sql = "INSERT INTO spf.cashflow(id, name, code, layeredattrs, direction, type, projectid, currency_id) VALUES ({0}, '{1}', '{2}', '{3}', '{4}', '{5}', {6}, {7});".format(
        id_cf_item, name, code, layeredattrs, direction, type, project_id, currency_id)
    sql_insert += sql + "\n"
    #print(sql)

    undo = "DELETE FROM spf.cashflow WHERE id={0};\n".format(id_cf_item) + undo

    id_cf_item += 1
    return id_cf_item

# Создаем "версию" плана CF
def create_cfversion(id, master_id, layeredattrs, versionid=1):
    global undo, sql_insert
    sql = "INSERT INTO spf.cashflowversion(id, layeredattrs, versionid, master_id) VALUES ({0}, '{1}', {2}, {3});".format(
        id, layeredattrs, versionid, master_id)
    sql_insert += sql + "\n"
    # print(sql)

    undo = "DELETE FROM spf.cashflowversion WHERE id={0};\n".format(id) + undo

    id += 1
    return id

def create_cat_item(id, parent_id, item_id, name, code, catalog_id=cashflowcatalog_id, versionid=1):
    global undo, sql_insert
    if (item_id):
        sql = "INSERT INTO spf.cashflowcatalogitem(id, versionid, catalog_id, parent_id, item_id) VALUES ({0}, {1}, {2}, {3}, {4});".format(
            id, versionid, catalog_id, parent_id, item_id)
    else:
        if (parent_id):
            sql = "INSERT INTO spf.cashflowcatalogitem(id, code, name, versionid, catalog_id, parent_id) VALUES ({0}, '{1}', '{2}', {3}, {4}, {5});".format(
                id, code, name, versionid, catalog_id, parent_id)
        else:
            sql = "INSERT INTO spf.cashflowcatalogitem(id, code, name, versionid, catalog_id) VALUES ({0}, '{1}', '{2}', {3}, {4});".format(
                id, code, name, versionid, catalog_id)
    sql_insert += sql + "\n"
    # print(sql)

    undo = "DELETE FROM spf.cashflowcatalogitem WHERE id={0};\n".format(id) + undo

    id += 1
    return id

def update_cf_item(id, layeredattrs, sql_update):
    sql_update = sql_update + "UPDATE spf.cashflow SET layeredattrs='{1}' WHERE id={0};".format(id, layeredattrs)
    return sql_update


rb = xlrd.open_workbook('data/indicators.xlsx')
sheet = rb.sheet_by_name("CashFlow")

for rownum in range(1, sheet.nrows):
    row = sheet.row_values(rownum)

    if (row[0] != "cf" and row[0] != "av" and row[0] != "ap" and row[0] != "cv" and row[0] != "cp" and row[0] != ""):
        num = row[0].split(";")

        if (len(num) == 1):
            cf_cat_item_id = create_cat_item(cf_cat_item_id, "", "", row[1], "cfci_" + str(cf_cat_item_id))
            parent_tag1 = cf_cat_item_id - 1
        if(len(num) == 2):
            cf_cat_item_id = create_cat_item(cf_cat_item_id, parent_tag1, "", row[1], "cfci_" + str(cf_cat_item_id))
            parent_tag2 = cf_cat_item_id - 1
        if (len(num) == 3):
            cf_cat_item_id = create_cat_item(cf_cat_item_id, parent_tag2, "", row[1], "cfci_" + str(cf_cat_item_id))
            parent_tag3 = cf_cat_item_id - 1
        if (len(num) == 4):
            cf_cat_item_id = create_cat_item(cf_cat_item_id, parent_tag3, "", row[1], "cfci_" + str(cf_cat_item_id))
        last_tag = cf_cat_item_id - 1

    if (row[0] == "cf"):
        fact = {}
        update_fact = {}
        plan = {}
        # За сколько периодов вносить данные
        for i in range(1, 20):
            key = periods[i]
            # До какого периода вносить фактические данные
            if(i<10):
                fact[key] = row[2 * (i - 1) + 5]
            # До какого периода вносить фактические данные (с помощью UPDATE)
            if(i<11):
                update_fact[key] = row[2 * (i - 1) + 5]
            plan[key] = row[2 * (i - 1) + 4]
            i += 1
        layeredattrs = json.dumps({"FACT": fact})
        layeredattrs_update_fact = json.dumps({"FACT": update_fact})
        layeredattrs_plan = json.dumps({"PLAN": plan})
        code = "cf_"+str(cf_item_id)
        cf_item_id = create_cf_item(cf_item_id, row[1], code, layeredattrs, row[2], row[3], currency_id)
        sql_update = update_cf_item(cf_item_id - 1, layeredattrs_update_fact, sql_update)
        cf_version_id = create_cfversion(cf_version_id, cf_item_id - 1, layeredattrs_plan)
        cf_cat_item_id = create_cat_item(cf_cat_item_id, last_tag, cf_version_id - 1, "", "")

#print (undo)
#print (sql_update)

undo_file = open('results/insert_cf_'+ datetime.now().strftime('%Y%m%d_%H%M') +'.txt', 'w')
undo_file.write(sql_insert)
undo_file.close()

undo_file = open('results/undo_cf_'+ datetime.now().strftime('%Y%m%d_%H%M') +'.txt', 'w')
undo_file.write(undo)
undo_file.close()

undo_file = open('results/update_cf_'+ datetime.now().strftime('%Y%m%d_%H%M') +'.txt', 'w')
undo_file.write(sql_update)
undo_file.close()

# В цвете вывод информации о завершении работы
OKGREEN = '\033[94m'
ENDC = '\033[0m'
print(OKGREEN  + "Скрипт {} закончил работу".format(sys.argv[0])+ ENDC)
