from unittest import TestCase
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


def get_cells(ws: Worksheet) -> list[list[any]]:
    result = []
    for row in ws.rows:
        new_row = []
        for cell in row:
            value = str(cell.value).strip()
            if value != '':
                new_row.append(value)
        result.append(new_row)
    return result


class TestNormalForm1(TestCase):
    def setUp(self):
        self.wb = load_workbook('tables1.xlsx')

    def test_table_has_correct_structure(self):
        self.assertIsNotNone(self.wb)
        sheet_names = self.wb.sheetnames
        self.assertEqual(len(sheet_names), 2)
        self.assertIn('persons', sheet_names)
        self.assertIn('grades', sheet_names)

    def test_table_has_correct_persons(self):
        cells = get_cells(self.wb['persons'])
        columns = [len(row) for row in cells]
        self.assertNotEqual(columns, [])
        self.assertTrue(all(map(lambda x: x == 3, columns)))
        values = [cell for row in cells for cell in row]
        self.assertGreaterEqual(values.count('Иванов'), 1)
        self.assertGreaterEqual(values.count('Иван'), 2)
        self.assertGreaterEqual(values.count('Иванович'), 1)
        self.assertGreaterEqual(values.count('Захаров'), 1)
        self.assertGreaterEqual(values.count('Егорович'), 1)
        self.assertGreaterEqual(values.count('Винокуров'), 1)
        self.assertGreaterEqual(values.count('Александр'), 1)
        self.assertGreaterEqual(values.count('Романович'), 1)

    def test_table_has_correct_grades(self):
        cells = [tuple(row) for row in get_cells(self.wb['grades'])]
        self.assertIn(('2', 'ФИИТ-22', 'Информатика', '5'), cells)
        self.assertIn(('2', 'ФИИТ-22', 'Математика', '4'), cells)
        self.assertIn(('2', 'ФИИТ-22', 'Базы данных', '4'), cells)
        self.assertIn(('5', 'ИВТ-22', 'Базы данных', '5'), cells)
        self.assertIn(('5', 'ИВТ-22', 'Математика', '4'), cells)


class TestNormalForm2(TestCase):
    def setUp(self):
        self.wb = load_workbook('tables2.xlsx')
