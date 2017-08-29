# coding: utf-8

import pandas as pd
from matplotlib import pyplot as plt
#import seaborn

df = pd.read_csv('data/gisp.csv')
df = df[['id','sum','period']]
print(df)


table = df.groupby('period').size()
print("Количество проектов, предоставленных ФРП: {} шт.".format(len(df)))

table.plot(kind='bar')
plt.title("Длительность проектов")
plt.xlabel("Длительность проекта, мес.")
plt.ylabel("Количество проектов, шт.")
plt.show()

table = df.groupby('sum').size()
table.plot(kind='bar')
plt.title("Объем запрашиваемых средств")
plt.xlabel("Запрашиваемые средства, тыс.руб.")
plt.ylabel("Количество проектов, шт.")
plt.show()


print("Минимальный объем запрашиваемых средств проектом: {} тыс. руб.".format(df[['sum']].min()))
exit()
print("Суммарный объем запрашиваемых средств в проектах ФРП: {} тыс. руб.".format(sum(summs)))
print("Средний запрашиваемый проектом объем финансирования: {} тыс. руб.".format(sum(summs)/len(summs)))

print("Максимальный объем запрашиваемых средств проектом: {} тыс. руб.".format(max(summs)))
