from unittest import TestCase
import spf_pi_v2


class TestProject(TestCase):
    def test_run_sql(self):
        #self.fail()
        pass

    def test_insert_pi_item(self):
        pass

    def test_insert_pi_version(self):
        pass

    def test_insert_pi_category_item(self):
        pass

    def test_update_pi_item(self):
        pass

    def test_import_pi_from_insert_file(self):
        a = spf_pi_v2.Project()
        a.import_pi_from_file('tests/indicators_etalon.xlsx')
        file = open('tests/insert_pi_etalon.txt', 'r')
        etalon = file.readlines()
        file.close()

        file = open(a.insert_result_file, 'r', encoding='utf-8')
        result = file.readlines()
        file.close()
        self.assertEqual(etalon, result)

    def test_import_pi_from_undo_file(self):
        a = spf_pi_v2.Project()
        a.import_pi_from_file('tests/indicators_etalon.xlsx')
        file = open('tests/undo_pi_etalon.txt', 'r')
        etalon = file.readlines()
        file.close()

        file = open(a.undo_result_file, 'r', encoding='utf-8')
        result = file.readlines()
        file.close()
        self.assertEqual(etalon, result)

    def test_import_pi_from_update_file(self):
        a = spf_pi_v2.Project()
        a.import_pi_from_file('tests/indicators_etalon.xlsx')
        file = open('tests/update_pi_etalon.txt', 'r')
        etalon = file.readlines()
        file.close()

        file = open(a.update_result_file, 'r', encoding='utf-8')
        result = file.readlines()
        file.close()
        self.assertEqual(etalon, result)
