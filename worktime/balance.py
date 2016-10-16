#!/usr/bin/env python
# -*- coding: utf-8 -*-

# --------------------------------------------------------------#
# Скрипт расчета управленческого баланса за период по ЦФУ ОСА   #
# --------------------------------------------------------------#

# Расчетный период в месяцах (с учетом декабря)
PERIOD = 10

import pandas as pd

## Функция обработки трудозатрат за месяц
def calc_month(i, writer):

    df = pd.read_excel("data/wt-0"+str(i)+".xls", sheetname=0, header=0, skiprows=7, skip_footer=1)

    pt = df.pivot_table(['Трудозатраты, ч'], ['Продукт'], aggfunc = 'sum', fill_value = 0)
    pt = pd.pivot_table(df, values = 'Трудозатраты, ч', index = ['Договор'], columns = ['Продукт'],  aggfunc = 'sum', fill_value = 0)
    try:
        pt.columns = ['Эксплуатация', 'Развитие', 'ExtTech', 'ДРТ']
    except:
        pt.columns = ['Эксплуатация', 'Развитие', 'ExtTech']
    pt['Всего'] = pt.sum(1)
    try:
        pt = pt[['Всего', 'Развитие', 'Эксплуатация', 'ExtTech', 'ДРТ']]
    except:
        pt = pt[['Всего', 'Развитие', 'Эксплуатация', 'ExtTech']]
    pt = pt.sort_values(['Всего', 'Развитие', 'Эксплуатация', 'ExtTech'], ascending = [0,0,0,0])

    #print(pt)
    pt.to_excel(writer,'№'+str(i))
    return pt

# Записать на лист в Excel
writer = pd.ExcelWriter('our.xlsx')

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

## Затраты на чужие ЦФУ
# Для каждого выгрузить список работ по чужим
# Преобразовать к виду "ЦФУ-время"
# Объединить все списки

## Аналогично нашим трудозатратам сделать за год чужие