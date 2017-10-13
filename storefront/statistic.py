# coding: utf-8

from bs4 import BeautifulSoup
import matplotlib.pyplot as plt
import os
import pandas as pd

dir = "I:/Projects/Инвестиционная площадка/Рабочие материалы/Проекты от ФРП/xml/"
files = os.listdir(dir)
xml_files = filter(lambda x: x.endswith('.xml'), files)

ids = []
summs =[]
months =[]

class project(object):
    def __init__(self, id, sum, month_return):
        self.id = id
        self.sum = sum # тыс.руб.
        self.month_return = month_return # мес.

for xml in xml_files:
    print(xml)

    with open(dir + xml, encoding="utf8") as fp:
        soup = BeautifulSoup(fp, 'lxml')

    #print(soup.prettify())

    prj = project(soup.html.body.nodes.arrayofprojectel.id.string.strip(), soup.html.body.nodes.arrayofprojectel.sum.string.strip(), soup.html.body.nodes.arrayofprojectel.month_return.string.strip())

    ids.append(int(prj.id))
    summs.append(int(prj.sum))
    months.append(int(prj.month_return))

#print(sorted(summs))
print("Количество проектов: {} шт.".format(len(summs)))
print("Суммарный объем запрашиваемых средств в проектах: {} тыс. руб.".format(sum(summs)))
print("Средний запрашиваемый проектом объем финансирования: {} тыс. руб.".format(sum(summs)/len(summs)))
print("Минимальный объем запрашиваемых средств проектом: {} тыс. руб.".format(min(summs)))
print("Максимальный объем запрашиваемых средств проектом: {} тыс. руб.".format(max(summs)))
#print(sorted(months))

d = {'id': ids, 'sum': summs, 'period':months}
df = pd.DataFrame(d, columns=['id','sum','period'])
df = df[['id','sum','period']]
df.to_csv('data/gisp.csv')