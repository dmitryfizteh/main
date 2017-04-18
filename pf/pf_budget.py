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
sheet = rb.sheet_by_index(3)

def sql_select(sql_str):
    result_sql = db.query(sql_str);
    return int(result_sql[0][0])

# Выбрать id_budjet_plan_item
title = "Подъездные автодороги и мосты"
id_budjet_plan_item = sql_select("SELECT id, title, description, order_num, budget_plan_id, parent_id FROM t_budget_plan_item WHERE title='{0}';".format(title))
print(id_budjet_plan_item)

expense_category = "Строительные работы" # Монтажные работы, Оборудование, Прочие затраты
expense_category_id = sql_select("SELECT id, name, description FROM t_expense_category WHERE name='{0}';".format(expense_category))
print(expense_category_id)


#agreement_stage_pmt_id
# Вставить в agreement_stage_expense
#db.query("INSERT INTO t_agreement_stage_expense(id, expense_category_id, expense_amount, budget_plan_item_id, agreement_stage_pmt_id) VALUES ({}, {}, {}, {},{});")
