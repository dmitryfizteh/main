from unittest import TestCase
import spf_pi_v2
import os


class TestProject(TestCase):
    def test_run_sql(self):
        #self.fail()
        a = spf_pi_v2.Project()
        a.insert_script = "Text1"
        a.undo_script = "Text2"
        a.run_sql("Add1", "Add2")
        self.assertEqual(a.insert_script, "Text1Add1")
        self.assertEqual(a.undo_script, "Add2Text2")

    def test_insert_pi_item(self):
        pass

    def test_insert_pi_version(self):
        pass

    def test_insert_pi_category_item(self):
        a = spf_pi_v2.Project()
        c = spf_pi_v2.PI_Category("Cat1", "cat_code1", "9", "2")
        a.insert_pi_category_item(c)
        self.assertEqual(a.insert_script, "INSERT INTO spf.projectindicatorcatalogitem(id, versionid, catalog_id, parent_id, item_id) " \
                     "VALUES ({0}, 1, 1, 2, 9);\n".format(a.max_pi_cat_item_id-1))
        a.insert_script = ""
        d = spf_pi_v2.PI_Category("Cat2", "cat_code2", "", "5")
        a.insert_pi_category_item(d)
        self.assertEqual(a.insert_script,
                         "INSERT INTO spf.projectindicatorcatalogitem(id, code, name, versionid, catalog_id, parent_id) " \
                         "VALUES ({0}, 'cat_code2', 'Cat2', 1, 1, 5);\n".format(a.max_pi_cat_item_id - 1))

    def test_update_pi_item(self):
        pass

    def test_import_pi_from_insert_file(self):
        a = spf_pi_v2.Project()
        a.connect('data/passwd.txt')
        a.fact_period = 9
        a.import_pi_from_file('tests/indicators_etalon.xlsx')
        file = open('tests/insert_pi_etalon.txt', 'r', encoding='utf-8')
        etalon = file.readlines()
        file.close()

        file = open(a.insert_result_file, 'r', encoding='utf-8')
        result = file.readlines()
        file.close()
        os.remove(a.insert_result_file)
        self.assertEqual(etalon, result)

    def test_import_pi_from_undo_file(self):
        a = spf_pi_v2.Project()
        a.connect('data/passwd.txt')
        a.fact_period = 9
        a.import_pi_from_file('tests/indicators_etalon.xlsx')
        file = open('tests/undo_pi_etalon.txt', 'r')
        etalon = file.readlines()
        file.close()

        file = open(a.undo_result_file, 'r', encoding='utf-8')
        result = file.readlines()
        file.close()
        os.remove(a.undo_result_file)
        self.assertEqual(etalon, result)

    def test_import_pi_from_update_file(self):
        a = spf_pi_v2.Project()
        a.connect('data/passwd.txt')
        a.fact_period = 9
        a.import_pi_from_file('tests/indicators_etalon.xlsx')
        file = open('tests/update_pi_etalon.txt', 'r')
        etalon = file.readlines()
        file.close()

        file = open(a.update_result_file, 'r', encoding='utf-8')
        result = file.readlines()
        file.close()
        os.remove(a.update_result_file)
        self.assertEqual(etalon, result)
