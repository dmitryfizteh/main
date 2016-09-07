#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd

# Выгрузить файл с бюджетом "Сис-админ" со статусом "Кроме удаленных" и датами периода
# Заменить суммы в валюте на рубли
# Проставить правильные статьи/договоры в незаполненных местах

df = pd.read_excel("data/ExportRunners.xls", sheetname=0, header=0, skiprows=0, skip_footer=0)
# TODO: Удалить строки с пустыми названиями
# Неточная заглушка для валют
df['Сумма'] = df[['Валюта', 'Сумма']].apply(lambda x: x[1] if (x[0] == 'Рубли') else x[1]*63, axis=1)
df['Без НДС'] = df[['Налог', 'Сумма']].apply(lambda x: x[1]/1.18 if (x[0] == 'С НДС') else x[1], axis=1)
df = df[['Название','Номер бегунка', 'Автор', 'Статья', 'Договор', 'Сумма', 'Налог',' Без НДС', 'Валюта', 'Дата создания', 'Описание']]
ptk = df.pivot_table(['Без НДС'], ['Статья'], aggfunc='sum', fill_value=0)
ptd = df.pivot_table(['Без НДС'], ['Статья', 'Договор'], aggfunc='sum', fill_value=0)
# TODO: добавить лимиты и равномерность расхода
# TODO: добавить общие суммы

writer = pd.ExcelWriter('budget.xlsx')
df.to_excel(writer, 'Выгрузка операций')
ptk.to_excel(writer, 'По-крупному')
ptd.to_excel(writer, 'Детально')
writer.save()