#!/usr/bin/env python
# -*- coding: utf-8 -*-

# --------------------------------------------------------------#
# Скрипт расчета управленческого баланса за период по ЦФУ ОСА   #
# --------------------------------------------------------------#

# Расчетный период в месяцах (с учетом декабря)
PERIOD = 6

import pandas as pd

df = pd.read_excel("data/2017/WorktimeReport.xlsx", sheetname="Первичка", header=0, skiprows=0, skip_footer=1)
df = df[['Period', 'Login', 'Fac', 'Agreement', 'Corrected']]

# Заменить строки с пустыми договорами на "пустой"
df['Agreement'] = df['Agreement'].fillna(value='Нет aggrement')

CFU_list = ['VerControl', 'СА-ExtTech', 'СА-развитие', 'СА-управление']
staff_list = ['ivanov', 'petrov', 'sidorov']
# Выбираем только свои ЦФУ
own_df = df[df['Fac'].isin(CFU_list)]
foreign_df = df[df['Login'].isin(staff_list)]
print(own_df.head())
print(foreign_df.head())

# Суммарно по ЦФУ
pt = own_df.pivot_table(['Corrected'], ['Fac'], aggfunc = 'sum', fill_value = 0)
# Сбрасываем MultiIndex
pt = pt.reset_index()
print(pt)

# TODO: сделать обработку пустых agreement
# По статьям
pt = own_df.pivot_table(values = 'Corrected', index = ['Agreement'], columns = ['Fac'],  aggfunc = 'sum', fill_value = 0)
pt = pt.reset_index()
pt['Всего'] = pt.sum(1)
pt = pt.sort_values(['Всего', 'СА-развитие', 'СА-управление'], ascending = [0,0,0])
summ = pd.DataFrame(pt.sum()).T
summ['Agreement'] = 'ИТОГО, часов:'
pt = pt.append(summ)

pt = pt[['Agreement', 'Всего', 'СА-развитие', 'СА-управление', 'VerControl']]

print(pt)

# В чужие ЦФУ
fpt = foreign_df.pivot_table(['Corrected'], ['Fac'],  aggfunc = 'sum', fill_value = 0)
print(fpt)

exit()

## Функция обработки трудозатрат за месяц
def calc_month(i, writer):
    df = pd.read_excel("data/2017/wt/wt-0"+str(i)+".xls", sheetname=0, header=0, skiprows=7, skip_footer=1)

    pt = df.pivot_table(['Трудозатраты, ч'], ['Продукт'], aggfunc = 'sum', fill_value = 0)

    pt = pd.pivot_table(df, values = 'Трудозатраты, ч', index = ['Договор'], columns = ['Продукт'],  aggfunc = 'sum', fill_value = 0)

    pt = pt.rename(columns={'Adm': 'Эксплуатация', 'Adm-Evol': 'Развитие', 'Adm-ExtTech': 'ExtTech', 'VerControl': 'ДРТ'})

    pt['Всего'] = pt.sum(1)

    pt = pt.sort_values(['Всего', 'Развитие', 'Эксплуатация'], ascending = [0,0,0])

    pt.to_excel(writer,'№'+str(i))
    return pt

# Записать на лист в Excel
writer = pd.ExcelWriter('balance2.xlsx')

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
