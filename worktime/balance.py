import pandas as pd
import requests

#http://plantime.office.custis.ru/ExportBugzillaManHours.aspx?id=367&fstart=01.08.2016&fend=31.08.2016

## Функция обработки трудозатрат за месяц
# Загрузить файл, загрузить его в dataframe

destination = 'all.xls'
url = 'http://plantime.office.custis.ru/ExportFinancialAccountingCenterBugzillaManHours.aspx?facs=39,38,84,153,162&fstart=01.08.2016&fend=31.08.2016'

user = input("Login:")
passw = input("Password:")
r = requests.get(url, stream=True, auth=(user, passw))
with open(destination, "wb") as code:
    code.write(r.content)

xl = pd.ExcelFile("all.xls")
df = xl.parse("Worktime")
df.head()

df = "FinancialAccountingCenterBugzillaManHours.xls"
# Построить матрицу, где по строкам виды деятельности, а по столбцам ЦФУ
# Записать на лист в Excel
# Записать dataframe в БД

## Обработка трудозатрат за год
# Запросами из БД построить матрицу, где по строкам виды деятельности, а по столбцам ЦФУ
# Записать на лист в Excel

## Количество рабочих дней за период, количество отпусков и больничных

## Затраты на чужие ЦФУ
# Для каждого выгрузить список работ по чужим
# Преобразовать к виду "ЦФУ-время"
# Объединить все списки

## Аналогично нашим трудозатратам сделать за год чужие