#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
count = 70
rb = xlrd.open_workbook('h:/bld.xlsx')
sheet = rb.sheet_by_index(0)
for rownum in range(sheet.nrows):
    row = sheet.row_values(rownum)
    if (row[0] != 0):
        for i in range(1,21):
            year = 11+i
            if (year > 24):
                year =year + 4
            str = "INSERT INTO t_cash_flow_item(id, title, item_sum, time_period_id, cash_flow_id) VALUES ({4}, '{0}', {3}, {2}, {1});".format(row[1],row[0], year, row[(i+2)], count)
            count = count + 1
            print(str)
