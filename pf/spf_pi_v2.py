# Script Name: spf_pi_v2.py
# Author: DmitryFizteh
# Created: 01/08/2017
# Version: v2
# Description: Скрипт наполнения данных Project indicators для СПФ

import xlrd
import json
import sys
from datetime import datetime


def save_script(filename, script):
    """Сохранение скрипта в файл"""
    # TODO: сделать сохранение в UTF-8
    file = open(filename + datetime.now().strftime('%Y%m%d_%H%M') + '.txt', 'w')
    file.write(script)
    file.close()


class PI(object):
    """Класс проектных индикаторов"""

    def __init__(self, name, code, values):
        self.name = name
        self.code = code
        self.values = values


class PI_Category(object):
    """Класс категорий проектных индикаторов"""

    def __init__(self, name, code, pi_version_id, parent_id):
        self.name = name
        self.code = code
        self.pi_version_id = pi_version_id
        self.parent_id = parent_id


class Project(object):
    """Мега-класс Проект"""

    def __init__(self):
        """В конструкторе задаются начальные параметры проекта"""
        self.undo_script = ""
        self.update_script = ""
        self.insert_script = ""
        self.project_import_file = 'data/indicators.xlsx'
        # TODO: считывать значения id из БД
        self.max_pi_item_id = 1000
        self.max_pi_cat_item_id = 5000
        self.max_pi_version_id = 1000
        # TODO: создавать проект
        self.project_id = 1
        # TODO: создавать периоды проекта
        # Перевод из периодов файла загрузки в периоды проекта в БД
        self.periods = {1: 8, 2: 9, 3: 11, 4: 12, 5: 13, 6: 14, 7: 15, 8: 16, 9: 17, 10: 18, 11: 19, 12: 20, 13: 21,
                        14: 22, 15: 23, 16: 24, 17: 25, 18: 26, 19: 27, 20: 28}

    def run_sql(self, insert_str, undo_str):
        """Наполнение скриптов"""
        self.insert_script = self.insert_script + insert_str
        self.undo_script = undo_str + self.undo_script

    def insert_pi_item(self, pi):
        insert_scr = "INSERT INTO spf.projectindicator(id, name, code, layeredattrs, projectid) " \
                     "VALUES ({0}, '{1}', '{2}', '{3}', '{4}');\n".format(
            self.max_pi_item_id, pi.name, pi.code, pi.values['fact'], self.project_id)
        undo_scr = "DELETE FROM spf.projectindicator WHERE id={0};\n".format(self.max_pi_item_id)
        self.run_sql(insert_scr, undo_scr)
        self.max_pi_item_id += 1

    def insert_pi_version(self, pi):
        # TODO: убрать ограничение на 1 версию
        insert_scr = "INSERT INTO spf.projectindicatorversion(id, layeredattrs, versionid, master_id) " \
                     "VALUES ({0}, '{1}', {2}, {3});\n".format(
            self.max_pi_version_id, pi.values['plan'], 1, self.max_pi_item_id - 1)
        undo_scr = "DELETE FROM spf.projectindicatorversion WHERE id={0};\n".format(self.max_pi_version_id)
        self.run_sql(insert_scr, undo_scr)
        self.max_pi_version_id += 1

    def insert_pi_category_item(self, category):
        # TODO: убрать ограничение на 1 версию и на 1 каталог
        if category.pi_version_id:
            insert_scr = "INSERT INTO spf.projectindicatorcatalogitem(id, versionid, catalog_id, parent_id, item_id) " \
                         "VALUES ({0}, {1}, {2}, {3}, {4});\n".format(
                self.max_pi_cat_item_id, 1, 1, category.parent_id, category.pi_version_id)
        else:
            if category.parent_id:
                insert_scr = "INSERT INTO spf.projectindicatorcatalogitem(id, code, name, versionid, catalog_id, parent_id) " \
                             "VALUES ({0}, '{1}', '{2}', {3}, {4}, {5});\n".format(
                    self.max_pi_cat_item_id, category.code, category.name, 1, 1, category.parent_id)
            else:
                insert_scr = "INSERT INTO spf.projectindicatorcatalogitem(id, code, name, versionid, catalog_id) " \
                             "VALUES ({0}, '{1}', '{2}', {3}, {4});\n".format(
                    self.max_pi_cat_item_id, category.code, category.name, 1, 1)
        undo_scr = "DELETE FROM spf.projectindicatorcatalogitem WHERE id={0};\n".format(self.max_pi_cat_item_id)
        self.run_sql(insert_scr, undo_scr)
        self.max_pi_cat_item_id += 1

    def update_pi_item(self, pi):
        self.update_script = self.update_script + "UPDATE spf.projectindicator SET layeredattrs='{1}' WHERE id={0};\n". \
            format(self.max_pi_item_id - 1, pi.values['update_fact'])

    def import_pi_from_file(self):
        rb = xlrd.open_workbook(self.project_import_file)
        sheet = rb.sheet_by_name("PI")

        for rownum in range(2, sheet.nrows):
            row = sheet.row_values(rownum)

            if row[0] != "":
                # TODO: поменять разделитель на точку (с учетом возможной запятой)
                num = str(row[0]).split(";")
                if len(num) == 1:
                    category = PI_Category(row[1], "pici_" + str(self.max_pi_cat_item_id), "", "")
                    self.insert_pi_category_item(category)
                    parent_tag1 = self.max_pi_cat_item_id - 1
                if len(num) == 2:
                    if parent_tag1 == "":
                        exit("ERROR: Неверная нумерация иерархии")
                    category = PI_Category(row[1], "pici_" + str(self.max_pi_cat_item_id), parent_tag1, last_tag)
                    self.insert_pi_category_item(category)
                    parent_tag2 = self.max_pi_cat_item_id - 1
                if len(num) == 3:
                    if parent_tag2 == "":
                        exit("ERROR: Неверная нумерация иерархии")
                    category = PI_Category(row[1], "pici_" + str(self.max_pi_cat_item_id), parent_tag2, last_tag)
                    self.insert_pi_category_item(category)
                    parent_tag3 = self.max_pi_cat_item_id - 1
                if len(num) == 4:
                    if parent_tag3 == "":
                        exit("ERROR: Неверная нумерация иерархии")
                    category = PI_Category(row[1], "pici_" + str(self.max_pi_cat_item_id), parent_tag3, last_tag)
                    self.insert_pi_category_item(category)
                last_tag = self.max_pi_cat_item_id - 1

            if row[0] == "":
                fact = {}
                update_fact = {}
                plan = {}
                for i in range(1, 20):
                    key = self.periods[i]
                    # До какого периода вносить фактические данные
                    if i < 10:
                        fact[key] = row[2 * (i - 1) + 4]
                    # До какого периода вносить фактические данные (с помощью UPDATE)
                    if i < 11:
                        update_fact[key] = row[2 * (i - 1) + 4]
                    plan[key] = row[2 * (i - 1) + 3]
                    i += 1

                values = {'plan': json.dumps({"PLAN": plan}), 'fact': json.dumps({"FACT": fact}),
                          'update_fact': json.dumps({"FACT": update_fact})}
                pi = PI(row[1], "pi_" + str(self.max_pi_item_id), values)
                if last_tag == "":
                    exit("ERROR: Неверная нумерация иерархии, CF item вне категорий")
                category = PI_Category(row[1], "", self.max_pi_version_id, last_tag)

                self.insert_pi_item(pi)
                self.insert_pi_version(pi)
                self.insert_pi_category_item(category)
                self.update_pi_item(pi)

        save_script('results/v2_insert_pi_', self.insert_script)
        save_script('results/v2_undo_pi_', self.undo_script)
        save_script('results/v2_update_pi_', self.update_script)


if __name__ == "__main__":
    P = Project()
    P.import_pi_from_file()
    del P

    # В цвете вывод информации о завершении работы
    OKGREEN = '\033[94m'
    ENDC = '\033[0m'
    print(OKGREEN + "Скрипт {} закончил работу".format(sys.argv[0]) + ENDC)
