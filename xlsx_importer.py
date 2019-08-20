"""
.py

Last Edited:

Lead Author[s]: Anthony Fong
Contributor[s]:

Description:


Machine I/O
Input:
Output:

User I/O
Input:
Output:


"""
########################################################################################################################

########## Libraries, Imports, & Setup ##########

# Default Libraries #
import shutil
import pathlib

# Downloaded Libraries #
import openpyxl
from openpyxl.utils import range_boundaries
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas
# import xlrd       need this one for pandas


# Local Libraries #


########## Definitions ##########

# Classes #
class xlsxDataFrame:
    def __init__(self, path=None, init=True):
        if path is None:
            self._path = None
        else:
            self.path = path
        self.op_workbook = None
        self.op_worksheets = dict()
        self.op_tables = dict()

        self.worksheets = dict()
        self._tables = dict()
        if init:
            self.load()

    def __getitem__(self, item):
        if isinstance(item, str):
            return self.worksheets[item]
        else:
            return self.op_workbook.worksheets[item]

    def __setitem__(self, key, value):
        pass

    @property
    def path(self):
        return self._path

    @path.setter
    def path(self, value):
        if isinstance(value, pathlib.Path) or value is None:
            self._path = value
        else:
            self._path = pathlib.Path(value)

    @property
    def sheet_names(self):
        return self.op_workbook.sheetnames

    @property
    def tables(self):
        return self._tables

    def load(self, in_path=None):
        self.load_wb(in_path=in_path)
        self.load_ws()
        self.load_tables()

    def load_wb(self, in_path=None):
        if in_path is not None:
            self.path = in_path
        self.op_workbook = openpyxl.load_workbook(self.path)

    def load_ws(self):
        for i, key in enumerate(self.sheet_names):
            self.op_worksheets[key] = self.op_workbook.worksheets[i]
            self.worksheets[key] = pandas.read_excel(self.path, index_col=None, header=None)

    def load_tables(self):
        op_tables = dict()
        tables = dict()
        for name in self.op_worksheets.keys():
            op_tables[name] = self.load_op_tables(name)
            tables[name] = self.load_pd_tables(name)
        self.op_tables = op_tables
        self._tables = tables

    def load_op_tables(self, name):
        tables = dict()
        for tbl in self.op_worksheets[name]._tables:
            tables[tbl.name] = tbl
        return tables

    def load_pd_tables(self, name):
        tables = dict()
        for tbl in self.op_worksheets[name]._tables:
            tables[tbl.name] = table2dataframe(self.op_worksheets[name], tbl)
        return tables

    def save(self):
        pass

    def copy_tables(self):
        pass








class xlxsImporter:
    def __init__(self, in_path=None, out_path=None):
        if in_path is None:
            self._path = None
        else:
            self.path = in_path
        if out_path is None:
            self._out_path = self.path
        else:
            self.out_path = out_path

        self.workbook = None

    @property
    def path(self):
        return self._path

    @path.setter
    def path(self, value):
        if isinstance(value, pathlib.Path) or value is None:
            self._path = value
        else:
            self._path = pathlib.Path(value)

    @property
    def out_path(self):
        return self._out_path

    @out_path.setter
    def out_path(self, value):
        if isinstance(value, pathlib.Path) or value is None:
            self._out_path = value
        else:
            self._out_path = pathlib.Path(value)















def table2dataframe(op_worksheet, table):
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)
    data = op_worksheet.iter_rows(min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col, values_only=True)
    headers = next(data)
    return pandas.DataFrame(data, columns=headers)


########## Main ##########
if __name__ == "__main__":
    db_path = pathlib.Path('C:/Users/ChangLab/Google Drive/Documents/Career/2017 - 2020 Chang Lab/Database')
    xfile = xlsxDataFrame(db_path.joinpath('SubjectNumbers.xlsx'))
    sheet = xfile[0]
    other = xfile.tables["Sheet1"]["SubjectTable"]

    print('done')
