#!/usr/bin/env python
# -*- coding: utf-8 -*-

# --------------------------------------------------------------#
# Скрипт расчета управленческого баланса за период по ЦФУ ОСА   #
# Version: 0.5                                                  #
# Author: Dmitry Morozov                                        #
# --------------------------------------------------------------#

# Выгрузить "трудозатраты по моим ЦФУ" за каждый месяц,
# и положить в папку <номер года>/wt/wt-<##>.xls

# Расчетный период в месяцах (с учетом декабря)
PERIOD = int(input("За какое количество месяцев (включая декабрь) рассчитываем: "))
year = "2017"

import pandas as pd

## Функция обработки трудозатрат за месяц
def calc_month(i, writer):
    if (i<10):
        df = pd.read_excel("data/"+year+"/wt/wt-0"+str(i)+".xls", sheetname=0, header=0, skiprows=7, skip_footer=1)
    else:
        df = pd.read_excel("data/"+year+"/wt/wt-" + str(i) + ".xls", sheetname=0, header=0, skiprows=7, skip_footer=1)

    pt = df.pivot_table(['Трудозатраты, ч'], ['Продукт'], aggfunc = 'sum', fill_value = 0)

    pt = pd.pivot_table(df, values = 'Трудозатраты, ч', index = ['Договор'], columns = ['Продукт'],  aggfunc = 'sum', fill_value = 0)

    pt = pt.rename(columns={'Adm': 'Эксплуатация', 'Adm-Evol': 'Развитие', 'Adm-ExtTech': 'ExtTech', 'VerControl': 'ДРТ'})

    pt['Всего'] = pt.sum(1)

    pt = pt.sort_values(['Всего', 'Развитие', 'Эксплуатация'], ascending = [0,0,0])

    pt.to_excel(writer,'№'+str(i))
    return pt

# Записать на лист в Excel
writer = pd.ExcelWriter('our_balance.xlsx')

## Обработка трудозатрат за год
# Построить матрицу, где по строкам виды деятельности, а по столбцам ЦФУ
frames = [calc_month(i, writer) for i in range(PERIOD)]
result = pd.concat(frames)
# Сбрасываем MultiIndex
result = result.reset_index()

result = pd.pivot_table(result, values = ['Всего', 'ExtTech', 'Развитие', 'Эксплуатация', 'ДРТ'], index = ['Договор'],  aggfunc = 'sum', fill_value = 0)
result = result.sort_values(['Всего', 'Развитие', 'Эксплуатация', 'ExtTech'], ascending = [0,0,0,0])
result = result.reset_index()
result = result[['Договор', 'Всего', 'Развитие', 'Эксплуатация', 'ExtTech', 'ДРТ']]

# ИТОГО
summ = pd.DataFrame(result.sum()).T
summ['Договор'] = 'ИТОГО, часов:'
result = result.append(summ)

print(result)
result.to_excel(writer,'За весь период')

writer.save()

exit()

## Количество рабочих дней за период, количество отпусков и больничных
