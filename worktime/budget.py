#!/usr/bin/env python
# -*- coding: utf-8 -*-

# --------------------------------------------------------------#
# Скрипт расчета трат бюджета за период по выгрузке из бегунков #
# --------------------------------------------------------------#

# Расчетный период в месяцах
PERIOD = 4

import pandas as pd
import xlsxwriter
### КАК ПОЛЬЗОВАТЬСЯ:
# Выставить PERIOD
# Выгрузить файл с бюджетом "Сис-админ" со статусом "Кроме удаленных" и датами периода
# Заменить суммы в валюте на рубли (в бегунках)
# Проставить правильные статьи/договоры в незаполненных местах (в бегунках)
# Временное: Проверить, что нет договоров с одинаковыми именами в разных статьях

### ВАЖНОЕ ПОВЕДЕНИЕ:
# Если договор бегунка не задан, то он попадет в расчет с договором "пустой". Расчет статей, договоров будет корректным.
# TODO: подсветить строки с пустым договором
# Все расчеты ведутся в рублях и без НДС.
# Статьи без лимитов попадают в расчет. TODO: подсветить строки без лимитов

# TODO: Выделить номера бегунков из гипепрссылок и загружать их
df = pd.read_excel("data/ExportRunners.xls", sheetname=0, header=0, skiprows=0, skip_footer=0)

# Удалить строки с пустыми названиями
df = df.dropna(subset=['Название'], how='all')
# Заменить строки с пустыми договорами на "пустой"
df['Договор'] = df['Договор'].fillna(value='пустой')

# Приводим все суммы в виду "Без НДС"
df['Без НДС'] = df[['Налог', 'Сумма']].apply(lambda x: x[1]/1.18 if (x[0] == 'С НДС') else x[1], axis=1)
# Оставляем только нужные столбцы
df = df[['Название', 'Автор', 'Статья', 'Договор', 'Сумма', 'Налог','Без НДС', 'Валюта', 'Дата создания', 'Описание']]
df_shot_k = df[['Статья', 'Без НДС']]
df_shot_d = df[['Статья', 'Договор', 'Без НДС']]
# Итого для таблицы трат
#summ_all = (df['Без НДС']).sum()

## ПО СТАТЬЯМ
# Загружаем лимиты
ptk_limits = pd.read_excel("data/limits.xlsx", sheetname='По-крупному', header=0, skiprows=0, skip_footer=0)
ptk_limits.columns = ['Статья','Лимит']
# Готовим строки таблицы лимитов для объединения со строками трат
ptk_shot = ptk_limits[['Статья']]
ptk_shot['Без НДС'] = 0
# Объединяем строки трат с нулевыми строками лимитов
ptk_for_pivot = pd.concat([df_shot_k, ptk_shot])
# Суммируем по статьям (по-крупному)
ptk = ptk_for_pivot.pivot_table(['Без НДС'], ['Статья'], aggfunc='sum', fill_value=0)
# Сбрасываем MultiIndex
ptk = ptk.reset_index()

# Cвязываем расходы и лимиты в один dataframe
ptk = pd.merge(ptk, ptk_limits, on='Статья', how='outer')
ptk['Лимит'] = ptk['Лимит'].fillna(value=0)
# Расчитываем остатки по статьям
ptk['Остаток'] = ptk[['Без НДС', 'Лимит']].apply(lambda x: x[1] - x[0], axis=1)
ptk['Равномерность расхода'] = ptk[['Без НДС', 'Лимит']].apply(lambda x: x[1]/12*PERIOD - x[0], axis=1)

# Рисуем "подвал"
summ_k = pd.DataFrame(ptk.sum()).T
summ_k['Статья'] = 'ИТОГО:'
ptk = ptk.append(summ_k)

## ДЕТАЛЬНО ПО ДОГОВОРАМ
# Загружаем лимиты
ptd_limits = pd.read_excel("data/limits.xlsx", sheetname='Детально', header=0, skiprows=0, skip_footer=0)
ptd_limits.columns = ['Статья', 'Договор', 'Лимит']
# Готовим строки таблицы лимитов для объединения со строками трат
ptd_shot = ptd_limits[['Статья', 'Договор']]
ptd_shot['Без НДС'] = 0
# Объединяем строки трат с нулевыми строками лимитов
ptd_for_pivot = pd.concat([df_shot_d, ptd_shot])

# Суммируем по договорам (детально)
ptd = ptd_for_pivot.pivot_table(['Без НДС'], ['Статья', 'Договор'], aggfunc='sum', fill_value=0)
# Сбрасываем MultiIndex
ptd = ptd.reset_index()

# Удаляем лишние столбцы
ptd_limits = ptd_limits[['Договор', 'Лимит']]

# Cвязываем расходы и лимиты в один dataframe
# TODO: Реализовать объединение не только по договору, но и по статье
ptd = pd.merge(ptd, ptd_limits, on='Договор',  how='outer')
ptd['Лимит'] = ptd['Лимит'].fillna(value=0)

# Расчитываем остатки по статьям
ptd['Остаток'] = ptd[['Без НДС', 'Лимит']].apply(lambda x: x[1] - x[0], axis=1)
ptd['Равномерность расхода'] = ptd[['Без НДС', 'Лимит']].apply(lambda x: x[1]/12*PERIOD - x[0], axis=1)

# Рисуем "подвал"
summ_d = pd.DataFrame(ptd.sum()).T
summ_d['Статья'] = 'ИТОГО:'
summ_d['Договор'] = ''
ptd = ptd.append(summ_d)

writer = pd.ExcelWriter('budget.xlsx', engine='xlsxwriter')
df.to_excel(writer, 'Выгрузка операций')
ptk.to_excel(writer, 'По-крупному')
ptd.to_excel(writer, 'Детально')

# TODO: добавить правильное форматирование Excel
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['По-крупному']
# Apply a conditional format to the cell range.
#  Add a format. Light red fill with dark red text.
format1 = workbook.add_format({'bg_color': '#FFC7CE','font_color': '#9C0006'})
#Write a conditional format over a range.
worksheet.conditional_format('E2:F10', {'type': 'cell',
                                         'criteria': '<',
                                         'value': 0,
                                         'format': format1})
writer.save()

# http://xlsxwriter.readthedocs.io/example_conditional_format.html


# TODO: реализовать прозрачный переброс средств между статьями
# TODO: реализовать тесты на попадание всех расходов и лимитов в общую сумму
