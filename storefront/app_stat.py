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
#plt.show()

table = df.groupby('sum').size()
table.plot(kind='bar')
plt.title("Объем запрашиваемых средств")
plt.xlabel("Запрашиваемые средства, тыс.руб.")
plt.ylabel("Количество проектов, шт.")
#plt.show()

df['В год'] = df[['sum', 'period']].apply(lambda x: x[0]/x[1]*12, axis=1)
table = df.groupby('В год').size()
table.plot(kind='bar')
plt.title("Объем запрашиваемых средств в год")
plt.xlabel("Запрашиваемые средства, тыс.руб.")
plt.ylabel("Количество проектов, шт.")
#plt.show()

df['gist'] = df[['sum']].apply(lambda x: 1 if x[0] > 50000 else 0, axis=1)
df['gist'] = df[['sum', 'gist']].apply(lambda x: 2 if x[0] > 100000 else x[1], axis=1)
df['gist'] = df[['sum', 'gist']].apply(lambda x: 3 if x[0] > 150000 else x[1], axis=1)
df['gist'] = df[['sum', 'gist']].apply(lambda x: 4 if x[0] > 200000 else x[1], axis=1)
df['gist'] = df[['sum', 'gist']].apply(lambda x: 5 if x[0] > 250000 else x[1], axis=1)
df['gist'] = df[['sum', 'gist']].apply(lambda x: 6 if x[0] > 300000 else x[1], axis=1)
df['gist'] = df[['sum', 'gist']].apply(lambda x: 7 if x[0] > 350000 else x[1], axis=1)
df['gist'] = df[['sum', 'gist']].apply(lambda x: 8 if x[0] > 400000 else x[1], axis=1)
df['gist'] = df[['sum', 'gist']].apply(lambda x: 9 if x[0] > 450000 else x[1], axis=1)
#print(df)
table = df.groupby('gist').size()
table.plot(kind='bar')
plt.title("Объем запрашиваемых средств в год")
plt.xlabel("Запрашиваемые средства, тыс.руб.")
plt.ylabel("Количество проектов, шт.")
plt.show()

print("Минимальный объем запрашиваемых средств проектом: {} тыс. руб.".format(df['sum'].min()))
print("Максимальный объем запрашиваемых средств проектом: {} тыс. руб.".format(df['sum'].max()))
print("Суммарный объем запрашиваемых средств в проектах: {} тыс. руб.".format(df['sum'].sum()))
print("Средний запрашиваемый проектом объем финансирования: {} тыс. руб.".format(df['sum'].sum()/len(df)))
print("Средний запрашиваемый проектом объем финансирования в год: {} тыс. руб.".format(df['В год'].sum()/len(df)))
exit()

