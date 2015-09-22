#!/usr/bin/env python
"""
    Extracts 'important' terms from raw input for processing by First_Level_Code assignment.
    Importance is determined by Part-Of-Speech Tagging via topia.termextract package.

    input: Arguments from the command line. First Argument is the file name. Next arguments are the names of the columns
    for processing.

    output: Multiple .xls files. File names are "*Column*.terms.xls".

    Nick Phillips - 2015
    phillips_nicholas@bah.com
"""

import sys
from xlrd import *

def get_sys_values(file_name):
    """
    Reads and returns names of surveillance system.
    :return: sys_names: list of names for systems.
    """
    current_book = open_workbook(file_name)
    active_sheet = current_book.sheet_by_index(0)
    sys_names = active_sheet.col_values(0)

    return sys_names

def get_col_values(col_name,file_name):

    current_book = open_workbook(file_name)
    active_sheet = current_book.sheet_by_index(0)
    row = active_sheet.row_values(0)
    index = 0
    for colTitle in row:
        if colTitle == col_name:
            col_data = active_sheet.col_values(index)
            return col_data
        index += 1

    print "Enter a valid column name!"
    return


#------------------------------------------------------

if __name__== "__main__":
    file_name = sys.argv[1]
    sys_values = get_sys_values(file_name)
    col_values = []
    for arg in sys.argv[2:]:
        col_values.append(get_col_values(arg,file_name))
