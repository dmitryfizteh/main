#!/usr/bin/env python
# -*- coding: utf-8 -*-

# ----------------------------------------------------------------#
# Скрипт расчета управленческого баланса за период по внешним ЦФУ #
# ----------------------------------------------------------------#

# Расчетный период в месяцах (с учетом декабря)
PERIOD = 4

import pandas as pd
import os

def calc_file(file_name):
    df = pd.read_excel(file_name, sheetname=0, header=0, skiprows=7, skip_footer=1)
    pt = pd.pivot_table(df, values = 'Трудозатраты, ч', index = ['Тема'],  aggfunc = 'sum', fill_value = 0)
    pt = pt.reset_index()
    # Удаляем ЦФУ ОСА, оставляем только внешние
    pt = pt[(pt['Тема'] != "VerControl") & (pt['Тема'] != "СА-развитие") & (pt['Тема'] != "СА-управление") & (pt['Тема'] != "СА-ExtTech")]
    return pt

def calc_month(month, year, writer):
    dir_string = "data/"+str(year)+"/"
    files = os.listdir(dir_string)
    if(month < 10):
        reg_string = '-0'+str(month)+'.xls'
    else:
        reg_string = '-' + str(month) + '.xls'
    current_files = filter(lambda x: x.endswith(reg_string), files)

    frames = [calc_file(dir_string+f) for f in current_files]
    result = pd.concat(frames)
    result = result.reset_index()

    result = pd.pivot_table(result, values = ['Трудозатраты, ч'], index = ['Тема'],  aggfunc = 'sum', fill_value = 0)
    result = result.reset_index()
    result.to_excel(writer, 'm'+str(month))
    return result

def calc_year(year, writer):
    months = [calc_month(m, year, writer) for m in range(PERIOD)]
    #months = [calc_month(m, year, writer) for m in range(8,10)]
    year_data = pd.concat(months)
    year_data = pd.pivot_table(year_data, values = ['Трудозатраты, ч'], index = ['Тема'],  aggfunc = 'sum', fill_value = 0)
    year_data = year_data.reset_index()
    #print(year_data)

    # ИТОГО
    summ = pd.DataFrame(year_data.sum()).T
    summ['Тема'] = 'ИТОГО, часов:'
    result = year_data.append(summ)

    # Удаляем ЦФУ ОСА, оставляем только внешние
    year_data_clear = year_data[(year_data['Тема']).str.contains("СА-управление")]
    #print(year_data_clear)

    summ = pd.DataFrame(year_data_clear.sum()).T
    summ['Тема'] = 'из них по инфра-ЦФУ, часов:'
    result = result.append(summ)

    result.to_excel(writer, 'За весь период')
    return year_data

writer = pd.ExcelWriter('foreign_balance.xlsx')
calc_year(2017, writer)
writer.save()
