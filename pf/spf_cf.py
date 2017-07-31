# Скрипт наполнения данных CashFlow для СПФ

#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import postgresql
import json

undo = ""

project_id = 1
#sql="INSERT INTO spf.project(id, name) VALUES (31,'Темповый проект');"
currency_id = 1

# TODO: Запросить id_cf_item (max_id)
id_cf_item = 5601
id_cf_version = 5601
id_cat_item = 5601

# TODO: Создать cashflowcatalog (max_id)
cashflowcatalog_id = 1

# TODO: Создать периоды проекта
periods = {1:27,2:23,3:9,4:32,5:30,6:21,7:7,8:18,9:15,10:12,11:4,12:28,13:26,14:24,15:10,16:33,17:31,18:22,19:8,20:19}

# Создаем CF и фактические значение
def create_cf_item(id_cf_item, name, code, layeredattrs, direction, type, currency_id):
    sql = "INSERT INTO spf.cashflow(id, name, code, layeredattrs, direction, type, projectid, currency_id) VALUES ({0}, '{1}', '{2}', '{3}', '{4}', '{5}', {6}, {7});".format(
        id_cf_item, name, code, layeredattrs, direction, type, project_id, currency_id)
    print(sql)

    global undo
    undo = "DELETE FROM spf.cashflow WHERE id={0};\n".format(id_cf_item) + undo

    id_cf_item += 1
    return id_cf_item

# Создаем "версию" плана CF
def create_cfversion(id, master_id, layeredattrs, versionid=1):
    sql = "INSERT INTO spf.cashflowversion(id, layeredattrs, versionid, master_id) VALUES ({0}, '{1}', {2}, {3});".format(
        id, layeredattrs, versionid, master_id)
    print(sql)

    global undo
    undo = "DELETE FROM spf.cashflowversion WHERE id={0};\n".format(id) + undo

    id += 1
    return id

def create_cat_item(id, parent_id, item_id, name, code, catalog_id=cashflowcatalog_id, versionid=1):

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
    print(sql)

    global undo
    undo = "DELETE FROM spf.cashflowcatalogitem WHERE id={0};\n".format(id) + undo

    id += 1
    return id

rb = xlrd.open_workbook('data/indicators.xlsx')
sheet = rb.sheet_by_name("CashFlow")

for rownum in range(1, sheet.nrows):
    row = sheet.row_values(rownum)

    if (row[0] != "cf" and row[0] != "av" and row[0] != "ap" and row[0] != "cv" and row[0] != "cp" and row[0] != ""):
        num = row[0].split(";")

        if (len(num) == 1):
            id_cat_item = create_cat_item(id_cat_item, "", "", row[1], "cfci_"+ str(id_cat_item) )
            parent_tag1 = id_cat_item-1
        if(len(num) == 2):
            id_cat_item = create_cat_item(id_cat_item, parent_tag1, "", row[1], "cfci_"+ str(id_cat_item))
            parent_tag2 = id_cat_item-1
        if (len(num) == 3):
            id_cat_item = create_cat_item(id_cat_item, parent_tag2, "", row[1], "cfci_"+ str(id_cat_item))
            parent_tag3 = id_cat_item-1
        if (len(num) == 4):
            id_cat_item = create_cat_item(id_cat_item, parent_tag3, "", row[1], "cfci_"+ str(id_cat_item))
        last_tag = id_cat_item-1

    if (row[0] == "cf"):
        fact = {}
        plan = {}
        for i in range(1, 20):
            key = periods[i]
            if(i<10):
                fact[key] = row[2 * (i - 1) + 5]
            plan[key] = row[2 * (i - 1) + 4]
            i += 1
        layeredattrs = json.dumps({"FACT": fact})
        layeredattrs_plan = json.dumps({"PLAN": plan})
        code = "cf_"+str(id_cf_item)
        id_cf_item = create_cf_item(id_cf_item, row[1], code, layeredattrs, row[2], row[3], currency_id)
        id_cf_version = create_cfversion(id_cf_version, id_cf_item-1, layeredattrs_plan)
        id_cat_item = create_cat_item(id_cat_item, last_tag, id_cf_version-1, "", "")

print (undo)