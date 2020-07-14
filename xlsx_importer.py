#!/usr/bin/env python
# -*- coding: utf-8 -*-
""" xlsx_importer.py
Description:
"""
__author__ = "Anthony Fong"
__copyright__ = "Copyright 2020, Anthony Fong"
__credits__ = ["Anthony Fong"]
__license__ = ""
__version__ = "1.0.0"
__maintainer__ = "Anthony Fong"
__email__ = ""
__status__ = "Prototype"

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


# Definitions #
# Classes #
class xlsxFile:
    """A class for reading and writing excel files.

    Args:
        path (str or :obj:'Path'): Path of excel file.
        init (bool, optional): If the object automatically populates itself.

    Attributes:


    """
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
        for tbl_name, tbl_range in self.op_worksheets[name].tables.items():
            tables[tbl_name] = tbl_range
        return tables

    def load_pd_tables(self, name):
        tables = dict()
        for tbl_name, tbl_range in self.op_worksheets[name].tables.items():
            tables[tbl_name] = range2dataframe(self.op_worksheets[name], tbl_range)
        return tables


# Functions#
def range2dataframe(op_worksheet, range_):
    """Converts a range of an excel sheet into a Pandas dataframe.

    Args:
        op_worksheet (:obj:'Worksheet'): The worksheet which the range is located.
        range_ (str): The range to convert into a dataframe. It must be in the excel standard for a range.

    Returns:
        :obj: DataFrame: The Pandas dataframe of the excel range.
    """
    min_col, min_row, max_col, max_row = range_boundaries(range_)
    data = op_worksheet.iter_rows(min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col, values_only=True)
    headers = next(data)
    return pandas.DataFrame(data, columns=headers)


# Main #
if __name__ == "__main__":
    db_path = pathlib.Path('/Users/changlab/Dropbox (UCSF Department of Neurological Surgery)/ChangLab/General Patient Info/EC223/')
    xfile = xlsxFile(db_path.joinpath('TrialNotes_EC223.xlsm'))
    sheet = xfile[0]
    other = xfile.tables["Tasks"]["PatientTasks"]

    print('done')

