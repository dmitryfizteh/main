#!/usr/bin/env python
# -*- coding: utf-8 -*-

# --------------------------------------------------------------#
# Скрипт расчета трат бюджета за период по выгрузке из бегунков #
# --------------------------------------------------------------#

# Расчетный период в месяцах
PERIOD = 9

import pandas as pd

# Выгрузить файл с бюджетом "Сис-админ" со статусом "Кроме удаленных" и датами периода
# Заменить суммы в валюте на рубли
# Проставить правильные статьи/договоры в незаполненных местах

df = pd.read_excel("data/ExportRunners.xls", sheetname=0, header=0, skiprows=0, skip_footer=0)
# TODO: Удалить строки с пустыми названиями
# TODO: Заменить строки с пустыми договорами на "резерв"
# Неточная заглушка для валют
df['Сумма'] = df[['Валюта', 'Сумма']].apply(lambda x: x[1] if (x[0] == 'Рубли') else x[1]*63, axis=1)
# Приводим все суммы в виду "Без НДС"
df['Без НДС'] = df[['Налог', 'Сумма']].apply(lambda x: x[1]/1.18 if (x[0] == 'С НДС') else x[1], axis=1)
# Оставляем только нужные столбцы
df = df[['Название','Номер бегунка', 'Автор', 'Статья', 'Договор', 'Сумма', 'Налог','Без НДС', 'Валюта', 'Дата создания', 'Описание']]

## ПО СТАТЬЯМ
# Суммируем по статьям (по-крупному)
ptk = df.pivot_table(['Без НДС'], ['Статья'], aggfunc='sum', fill_value=0)
# Загружаем лимиты
ptk_limits = pd.read_excel("data/limits.xlsx", sheetname='По-крупному', header=0, skiprows=0, skip_footer=0)
ptk_limits.columns = ['Статья','Лимит']
# Сбрасываем MultiIndex
ptk = ptk.reset_index()
# Cвязываем расходы и лимиты в один dataframe
ptk = pd.merge(ptk, ptk_limits, left_on='Статья', right_on='Статья')
# Расчитываем остатки по статьям
ptk['Остаток'] = ptk[['Без НДС', 'Лимит']].apply(lambda x: x[1] - x[0], axis=1)
ptk['Равномерность расхода'] = ptk[['Без НДС', 'Лимит']].apply(lambda x: x[1]/12*PERIOD - x[0], axis=1)

## ДЕТАЛЬНО ПО ДОГОВОРАМ
# Суммируем по договорам (детально)
ptd = df.pivot_table(['Без НДС'], ['Статья', 'Договор'], aggfunc='sum', fill_value=0)
# Загружаем лимиты
ptd_limits = pd.read_excel("data/limits.xlsx", sheetname='Детально', header=0, skiprows=0, skip_footer=0)
ptd_limits.columns = ['Статья', 'Договор', 'Лимит']
# Удаляем лишние столбцы
ptd_limits = ptd_limits[['Договор', 'Лимит']]

# TODO: Cвязать расходы и лимиты в один dataframe (пока заглушка)
ptd['Лимит'] = 0
# Расчитываем остатки по статьям
ptd['Остаток'] = ptd[['Без НДС', 'Лимит']].apply(lambda x: x[1] - x[0], axis=1)
ptd['Равномерность расхода'] = ptd[['Без НДС', 'Лимит']].apply(lambda x: x[1]/12*PERIOD - x[0], axis=1)

# TODO: добавить общие суммы

writer = pd.ExcelWriter('budget.xlsx')
df.to_excel(writer, 'Выгрузка операций')
ptk.to_excel(writer, 'По-крупному')
ptd.to_excel(writer, 'Детально')
# TODO: добавить правильное форматирование Excel
writer.save()
