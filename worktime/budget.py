#!/usr/bin/env python
# -*- coding: utf-8 -*-

# --------------------------------------------------------------#
# Скрипт расчета трат бюджета за период по выгрузке из бегунков #
# --------------------------------------------------------------#

# Расчетный период в месяцах
PERIOD = 8
# номера бегунков (с прошлого периода) - не учитывать их
exception_list_2017 = [16667, 16705, 16706, 16714, 16726, 16726, 16743,16768, 16769, 16836, 17594, 17739, 17592]

import pandas as pd
import xlsxwriter

# Выставить PERIOD
# Выгрузить файл с бюджетом "Сис-админ" со статусом "Кроме удаленных" и датами периода
# Заменить форматирование столбца с номерами бегунков на "текстовый" TODO: делать автоматически и проверять что нет прошлогодних
# Заменить суммы в валюте на рубли (в бегунках)
# Проставить правильные статьи/договоры в незаполненных местах (в бегунках)
# Временное (до todo4): Проверить, что нет договоров с одинаковыми именами в разных статьях

df = pd.read_excel("data/ExportRunners.xls", sheetname=0, header=0, skiprows=0, skip_footer=0)

#Преобразовать ссылки  в номера бегунков

#print(df.info())
#print(df.head())
#exit()

# сделать исключения номеров бегунков (с прошлого периода) - НЕ РАБОТАЮТ, пока нет перевода ссылок в номера бегунков
df = df[~df.isin({'Номер бегунка':exception_list_2017})]
df = df.dropna(subset=['Номер бегунка'], how='all')



# Удалить строки с пустыми названиями
df = df.dropna(subset=['Название'], how='all')
# Заменить строки с пустыми договорами на "пустой"
df['Договор'] = df['Договор'].fillna(value='пустой')

#Если есть незаполненные договоры, то вывести предупреждение
df_error = df[df['Договор'].str.contains("пустой")]
if (df_error['Статья'].count() > 0):
    print("ОШИБКА!!! Пустые договоры в файле!")
    print(df_error)

# Приводим все суммы в виду "Без НДС"
df['Без НДС'] = df[['Налог', 'Сумма']].apply(lambda x: x[1]/1.18 if (x[0] == 'С НДС') else x[1], axis=1)
# Оставляем только нужные столбцы
df = df[['Название','Автор', 'Статья', 'Договор', 'Сумма', 'Налог','Без НДС', 'Валюта', 'Дата создания', 'Описание']]
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
ptk.loc[(ptk['Статья'].str.contains("резерв")), 'Равномерность расхода'] = 0;

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
# TODO:4 Реализовать объединение не только по договору, но и по статье (чтобы можно было заводить договоры с одинаковыми именами в разных статьях)
ptd = pd.merge(ptd, ptd_limits, on='Договор',  how='outer')
ptd['Лимит'] = ptd['Лимит'].fillna(value=0)

# Расчитываем остатки по статьям
ptd['Остаток'] = ptd[['Без НДС', 'Лимит']].apply(lambda x: x[1] - x[0], axis=1)
ptd['Равномерность расхода'] = ptd[['Без НДС', 'Лимит','Статья']].apply(lambda x: x[1]/12*PERIOD - x[0], axis=1)
ptd.loc[(ptd['Статья'].str.contains("резерв")), 'Равномерность расхода'] = 0;

# Рисуем "подвал"
summ_d = pd.DataFrame(ptd.sum()).T
summ_d['Статья'] = 'ИТОГО:'
summ_d['Договор'] = ''
ptd = ptd.append(summ_d)

writer = pd.ExcelWriter('budget.xlsx', engine='xlsxwriter')
df.to_excel(writer, 'Выгрузка операций')
ptk.to_excel(writer, 'По-крупному')
ptd.to_excel(writer, 'Детально')

# TODO:5 добавить правильное форматирование Excel (финансовое форматирование, уменьшенный шрифт, выравнивание ячеек)
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
format1 = workbook.add_format({'bg_color': '#FFC7CE','font_color': '#9C0006'})
money_fmt = workbook.add_format({'num_format': '#,##0'})

worksheet = writer.sheets['По-крупному']
worksheet.conditional_format('E2:F10', {'type': 'cell',
                                         'criteria': '<',
                                         'value': 0,
                                         'format': format1})
worksheet.set_column('C:F', 12, money_fmt)
# Выравнивание
worksheet.set_column('B:B', 19)
worksheet.set_column('F:F', 22, money_fmt)
worksheet.set_zoom(90)

# пометить красным, если отрицательное значение
worksheet = writer.sheets['Детально']
worksheet.conditional_format('F2:F60', {'type': 'cell',
                                         'criteria': '<',
                                         'value': 0,
                                         'format': format1})
# пометить красным, если "пустой" договор
worksheet.conditional_format('C2:C60', {'type': 'bottom',
                                         'value': 'пустой',
                                         'format': format1})
# Add a number format for cells with money.
worksheet.set_column('D:G', 12, money_fmt)

# Выравнивание
worksheet.set_column('B:B', 19)
worksheet.set_column('C:C', 36)
worksheet.set_column('G:G', 22)
worksheet.set_zoom(90)

worksheet = writer.sheets['Выгрузка операций']
worksheet.set_column('B:K', 22)
worksheet.set_zoom(90)
writer.save()

# http://xlsxwriter.readthedocs.io/example_conditional_format.html


# TODO:6 реализовать прозрачный переброс средств между статьями (через версии бюджетов)
# TODO:7 реализовать тесты на попадание всех расходов и лимитов в общую сумму
