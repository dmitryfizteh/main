# Скрипт наполнения данных CashFlow для СПФ

#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import postgresql

undo = ""

project_id = 1
currency_id = 1

# TODO: Создать cashflowcatalog (max_id)
cashflowcatalog_id = 1

# TODO: Создать периоды проекта

# Создаем CF
def create_cf_item(id_cf_item, name, code, layeredattrs, direction, currency_id):
    str = "INSERT INTO cashflow(id, name, code, layeredattrs, direction, currency_id) VALUES ({0}, '{1}', {2}, {3}, {4}, {5});".format(
        id_cf_item, name, code, layeredattrs, direction, currency_id)
    print(str)
    global undo
    undo = "DELETE FROM cashflow WHERE id={0};\n".format(id_cf_item) + undo

    return id_cf_item

rb = xlrd.open_workbook('h:/indicators.xlsx')
sheet = rb.sheet_by_name("CashFlow")

# TODO: сделать обработку всего файла
#for rownum in range(2, sheet.nrows):
for rownum in range(3, 4):
    row = sheet.row_values(rownum)
print(row)



d = {"1": 883.24, "2": 283.06, "3": 831.27}
layeredattrs = {"FACT": d}

print(layeredattrs)

# {"FACT": {"1": 883.24, "2": 283.06, "3": 831.27, "4": 614.49, "5": 466.86, "6": 813.58, "7": 847.83, "8": 367.32, "9": 975.97, "10": 529.49, "11": 942.07, "12": 874.22, "13": 963.54, "14": 844.62, "15": 625.83, "16": 295.44, "17": 869.70, "18": 931.22, "19": 418.13, "20": 236.79, "21": 773.25, "22": 200.03, "23": 496.35, "24": 944.35, "25": 882.83, "26": 111.54, "27": 466.46, "28": 231.67, "29": 890.41, "30": 891.07, "31": 616.47, "32": 753.94, "33": 178.49}}