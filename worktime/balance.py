#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd
import requests

# Загрузить файл, загрузить его в dataframe
'''
destination = 'all.xls'
url = 'http://'
user = input("Login:")
passw = input("Password:")
r = requests.get(url, stream=True, auth=(user, passw))
with open(destination, "wb") as code:
    code.write(r.content)
'''

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
    pt.to_excel(writer,'Sheet'+str(i))
    return pt

# Записать на лист в Excel
writer = pd.ExcelWriter('our.xlsx')

frames = [calc_month(i, writer) for i in range(9)]
result = pd.concat(frames)
print(result.head())
exit()
result = pd.pivot_table(result, values = ['Всего', 'ExtTech', 'Развитие', 'Эксплуатация', 'ДРТ'], index = ['Договор'],  aggfunc = 'sum', fill_value = 0)

writer.save()

print(result)

exit()

## Количество рабочих дней за период, количество отпусков и больничных

## Затраты на чужие ЦФУ
# Для каждого выгрузить список работ по чужим
# Преобразовать к виду "ЦФУ-время"
# Объединить все списки

## Аналогично нашим трудозатратам сделать за год чужие