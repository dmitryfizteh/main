#!/usr/bin/env python
# -*- coding: utf-8 -*-

# ----------------------------------------------------------------#
# Скрипт расчета управленческого баланса за период по внешним ЦФУ #
# ----------------------------------------------------------------#

# Расчетный период в месяцах (с учетом декабря)
PERIOD = 10

import pandas as pd
df = pd.read_excel("data/09/Morozov-09.xls", sheetname=0, header=0, skiprows=7, skip_footer=1)
pt = pd.pivot_table(df, values = 'Трудозатраты, ч', index = ['Тема'],  aggfunc = 'sum', fill_value = 0)
pt = pt.reset_index()

# Удаляем ЦФУ ОСА, оставляем только внешние
pt = pt[pt['Тема'] != "VerControl"]
pt = pt[pt['Тема'] != "СА-развитие"]
pt = pt[pt['Тема'] != "СА-управление"]

print(pt)
