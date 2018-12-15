import unittest
import os
import xlrd

from excel_factory import ExcelFactory, ExcelFactory2


class TestExcelFactory(unittest.TestCase):

    def test_build_simple_excel_file_1(self):
        factory = ExcelFactory()
        excel_file = "myexcel1.xlsx"
        sheet_name = "sheetX"
        header_list = ['header1', 'header2', 'header3', 'header4']
        factory.create_workbook(excel_file) \
            .add_sheet(sheet_name) \
            .add_header(header_list) \
            .add_values_to_column('header1', ['4711', '9999']) \
            .add_values_to_column('header2', ['0.45', '0.77']) \
            .add_values_to_column('header4', ['10000']) \
            .build()

        self.assertIsNotNone(os.path.isfile(excel_file))
        workbook = xlrd.open_workbook(excel_file)
        self.assertIsNotNone(workbook)
        worksheet = workbook.sheet_by_name(sheet_name)
        self.assertIsNotNone(worksheet)
        col_index = 0
        for col in header_list:
            self.assertEqual(col, worksheet.cell(0, col_index).value, "No match for columns %s" % col)
            col_index += 1
        self.assertEqual('4711', worksheet.cell(1, 0).value)
        self.assertEqual('9999', worksheet.cell(2, 0).value)
        self.assertEqual('0.45', worksheet.cell(1, 1).value)
        self.assertEqual('0.77', worksheet.cell(2, 1).value)
        self.assertEqual('10000', worksheet.cell(1, 3).value)
        os.remove(excel_file)

    def test_build_simple_excel_file_2(self):
        # create file
        factory = ExcelFactory2()
        excel_file = "myexcel2.xlsx"
        sheet_name = "sheetX"
        factory.create_workbook(excel_file)
        factory.add_sheet(sheet_name)
        header = ['header_col1', 'header_col2', 'header_col3', 'header_col4', 'header_col5']
        row1 = ['row_1_value', 'row_1_value', 'row_1_value', 'row_1_value', 'row_1_value']
        row2 = ['row_2_value', 'row_2_value', 'row_2_value', 'row_2_value', 'row_2_value']
        factory.add_rows([header, row1, row2])
        # test file
        self.assertIsNotNone(os.path.isfile(excel_file))
        workbook = xlrd.open_workbook(excel_file)
        self.assertIsNotNone(workbook)
        worksheet = workbook.sheet_by_name(sheet_name)
        self.assertIsNotNone(worksheet)
        self.check_values_in_row(0, header, worksheet)
        self.check_values_in_row(1, row1, worksheet)
        self.check_values_in_row(2, row2, worksheet)
        os.remove(excel_file)

    def check_values_in_row(self, rowx, values, worksheet):
        col_index = 0
        for value in values:
            self.assertEqual(value, worksheet.cell(rowx, col_index).value, "No match for columns %s" % value)
            col_index += 1
