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
import re
import csv


def get_sys_values(file_name):
    """
    Reads and returns names of surveillance system.
    :return: sys_names: list of names for systems.
    """
    current_book = open_workbook(file_name)
    active_sheet = current_book.sheet_by_index(0)
    sys_names = active_sheet.col_values(0)

    return sys_names


def get_col_values(col_name, file_name):
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

def extract_terms(phrase,key_terms):
    """
    Initially intended to use NLP (POS recognition) for term extraction. However, as many terms are not
    structured like text, better extraction is achieved by extracting all words and removing stop words with
    NLTK.
    :param phrase:
    :return: list of non-stop words in phrase
    """

    temp_terms = []

    pattern = re.compile('[\W_]+')
    phrase = pattern.sub(" ",phrase)
    phrase = phrase.lower()
    for term in key_terms:
        condition = True
        for i in term:
            if i not in phrase:
                condition = False
        if condition:
            temp_terms.append(" ".join(term))

    temp_check = " ".join(temp_terms)
    #Perform text filtering here.
    stop = stopwords.words('english')
    temp = [i for i in phrase.split() if i not in stop]
    temp = [i for i in temp if not i.isdigit()]
    temp = [i for i in temp if len(i) > 2]
    temp = [i for i in temp if i not in temp_check]
    temp_terms = temp_terms + temp

    return temp_terms


#------------------------------------------------------

if __name__ == "__main__":
    file_name = sys.argv[1]
    sys_values = get_sys_values(file_name)
    col_values = []
    for arg in sys.argv[2:]:
        col_values.append(get_col_values(arg,file_name))

    key_terms = get_sys_values("healthmap_codes.xls")
    key_terms = [i.split() for i in key_terms]

    processed_data = []
    processed_single_col = []
    for single_col in col_values:
        for single_phrase in single_col:
            processed_single_col.append(extract_terms(single_phrase, key_terms))
        processed_data.append(processed_single_col)
        processed_single_col = []

    for new_book_data in processed_data:
        with open(file_name[:len(file_name)-5]+"."+new_book_data[0][0]+".csv",'wb') as csvfile:
            workbookOut = csv.writer(csvfile)
            for row in new_book_data[1:len(new_book_data)]:
                workbookOut.writerow(row)