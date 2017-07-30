# Скрипт наполнения данных Project indicators для СПФ

#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import postgresql
import json

undo = ""

projectid = 1
#sql="INSERT INTO spf.project(id, name) VALUES (31,'Темповый проект');"

id_pi_item = 1000
pi_catalog_id = 1
id_cat_item = 5000
id_pi_version = 1000

# Создаем PI и фактические значение
def create_pi_item(id_pi_item, name, code, layeredattrs, projectid):
    sql = "INSERT INTO spf.projectindicator(id, name, code, layeredattrs, projectid) VALUES ({0}, '{1}', '{2}', '{3}', '{4}');".format(
        id_pi_item, name, code, layeredattrs, projectid)
    print(sql)

    global undo
    undo = "DELETE FROM spf.projectindicator WHERE id={0};\n".format(id_pi_item) + undo

    id_pi_item += 1
    return id_pi_item

# Создаем "версию" плана PI
def create_piversion(id, master_id, layeredattrs, versionid=1):
    sql = "INSERT INTO spf.projectindicatorversion(id, layeredattrs, versionid, master_id) VALUES ({0}, '{1}', {2}, {3});".format(
        id, layeredattrs, versionid, master_id)
    print(sql)

    global undo
    undo = "DELETE FROM spf.projectindicatorversion WHERE id={0};\n".format(id) + undo

    id += 1
    return id

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
    print(sql)

    global undo
    undo = "DELETE FROM spf.projectindicatorcatalogitem WHERE id={0};\n".format(id) + undo

    id += 1
    return id

rb = xlrd.open_workbook('data/indicators.xlsx')
sheet = rb.sheet_by_name("PI")

# TODO: сделать обработку всего файла
#for rownum in range(2, sheet.nrows):
for rownum in range(2, 5):
    row = sheet.row_values(rownum)

    if (row[0] != ""):
        num = str(row[0]).split(";")

        if (len(num) == 1):
            id_cat_item = create_cat_item(id_cat_item, "", "", row[1], "pici_"+ str(id_cat_item) )
            parent_tag1 = id_cat_item-1
        if(len(num) == 2):
            id_cat_item = create_cat_item(id_cat_item, parent_tag1, "", row[1], "pici_"+ str(id_cat_item))
            parent_tag2 = id_cat_item-1
        if (len(num) == 3):
            id_cat_item = create_cat_item(id_cat_item, parent_tag2, "", row[1], "pici_"+ str(id_cat_item))
            parent_tag3 = id_cat_item-1
        if (len(num) == 4):
            id_cat_item = create_cat_item(id_cat_item, parent_tag3, "", row[1], "pici_"+ str(id_cat_item))
        last_tag = id_cat_item-1

    if (row[0] == ""):
        fact = {}
        plan = {}
        for i in range(1, 20):
            key="{}".format(i)
            fact[key] = row[2 * (i - 1) + 4]
            plan[key] = row[2 * (i - 1) + 3]
            i += 1
        layeredattrs = json.dumps({"FACT": fact})
        layeredattrs_plan = json.dumps({"PLAN": plan})
        code = "pi_"+str(id_pi_item)
        id_pi_item = create_pi_item(id_pi_item, row[1], code, layeredattrs, projectid)
        id_pi_version = create_piversion(id_pi_version, id_pi_item-1, layeredattrs_plan)
        id_cat_item = create_cat_item(id_cat_item, last_tag, id_pi_version-1, "", "")

print (undo)