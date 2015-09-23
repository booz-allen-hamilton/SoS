#!/usr/bin/env python
"""
    Extracts information from master list, formats codes for first level processing by First_Level_Code.py.

    input: Arguments from the command line. First Argument is the file name. Next arguments are the names of the columns
    for processing.

    output: Multiple .xls files. File names are "*Column*.terms.xls".

    Nick Phillips - 2015
    phillips_nicholas@bah.com
"""

import sys
from xlrd import *
from topia.termextract import tag
from topia.termextract import extract
from nltk.corpus import stopwords
from xlwt import *


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
    """
    Gets values for a desired column from the master file
    :param col_name: column of interest
    :param file_name: name of master file
    :return: List of column data, if found. Nothing otherwise.
    """
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

def extract_terms(phrase):
    """
    Initially intended to use NLP (POS recognition) for term extraction. However, as many terms are not
    structured like text, better extraction is achieved by extracting all words and removing stop words with
    NLTK.
    :param phrase:
    :return: list of non-stop words in phrase
    """
    stop = stopwords.words('english')
    temp = [i for i in phrase.split() if i not in stop]
    return [i for i in temp if not has_numbers(i)]


def has_numbers(input_string):
    return any(char.isdigit() for char in input_string)


#------------------------------------------------------

if __name__ == "__main__":
    file_name = sys.argv[1]
    sys_values = get_sys_values(file_name)
    col_values = []
    for arg in sys.argv[2:]:
        col_values.append(get_col_values(arg,file_name))

    processed_data = []
    processed_single_col=[]
    for single_col in col_values:
        for single_phrase in single_col:
            processed_single_col.append(extract_terms(single_phrase))
        processed_data.append(processed_single_col)
        processed_single_col = []

    col_val_index = 0
    for new_book_data in processed_data:
        workbookOut = Workbook()
        sheet1 = workbookOut.add_sheet("Extracted Terms")

        row_c = 0
        col_c = 2
        for sys_terms in new_book_data[1:len(new_book_data)]:
            sheet1.write(row_c, 0, sys_values[row_c+1])
            for w in sys_terms:
                sheet1.write(row_c, col_c, w)
                col_c += 1
            col_c = 2
            row_c += 1

        row_c = 0
        for col in col_values[col_val_index][1:len(col_values[col_val_index])]:
            sheet1.write(row_c, 1, col_values[col_val_index][row_c+1])
            row_c += 1

        workbookOut.save(file_name[:len(file_name)-5]+"."+new_book_data[0][0]+".xls")
        col_val_index += 1