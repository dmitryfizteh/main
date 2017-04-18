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

f = open('passwd.txt')
for connect_str in f:
    db = postgresql.open(connect_str)
f.close()

rb = xlrd.open_workbook('h:/indicators.xlsx')
sheet = rb.sheet_by_index(2)

id_wbs_sql = db.query("SELECT MAX(id) FROM t_wbs_item;");
id_wbs = int (id_wbs_sql[0][0]) + 1

def create_wbs(id_wbs, name, start_date, end_date, project_id, parent_id):
   # if( parent_id == ""):
   #     print("INSERT INTO t_wbs_item (id, title, start_date_planned, end_date_planned, project_id) VALUES ({0}, '{1}', '{2}', '{3}', {4});".format(id_wbs, name, start_date, end_date, project_id))
   # else:
    print("INSERT INTO t_wbs_item (id, title, start_date_planned, end_date_planned, project_id, parent_id) VALUES ({0}, '{1}', '{2}', '{3}', {4}, {5});".format(id_wbs, name, start_date, end_date, project_id, parent_id))

    global undo
    undo = "DELETE FROM t_wbs_item WHERE id={0};\n".format(id_wbs) + undo

for rownum in range(2, sheet.nrows):
    row = sheet.row_values(rownum)
    #print(row)

    num = str(row[0]).split(".")
    if (len(num) == 1):
        create_wbs(id_wbs, row[1], row[2], row[3], project_id, "1010") # Код корневого каталога
        parent_id1 = id_wbs
        id_wbs += 1
    if (len(num) == 2):
        create_wbs(id_wbs, row[1], row[2], row[3], project_id, parent_id1)
        parent_id2 = id_wbs
        id_wbs += 1
    if (len(num) == 3):
        create_wbs(id_wbs, row[1], row[2], row[3], project_id, parent_id2)
        parent_id3 = id_wbs
        id_wbs += 1
    if (len(num) == 4):
        create_wbs(id_wbs, row[1], row[2], row[3], project_id, parent_id3)
        id_wbs += 1

undo_file = open('undo_wbs_'+ datetime.now().strftime('%Y%m%d_%H%M%S') +'.txt', 'w')
undo_file.write(undo)
undo_file.close()