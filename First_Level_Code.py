#!/usr/bin/env python
"""
First level code assignment script for data in SoS project.

input: .xlsx file given as first argument in command line. Column "A" and "B" are surveillance system and condition
 information, respectively. Columns "C" through last are individual key terms from condition information column.

output: Assigns "A" level codes starting at "A1" to each surveillance system. Codes are ordered as they appear in the
input file. Each unique word receives a code. Codes are written to "*filename*.out.xls" with two sheets. Sheet
"System Codes" assigns codes to systems. Sheet "Code Legend" describes code legend.

Example:
input: python First_Level_Code.py Data_Workbook.xlsx
output: Data_Workbook.out.xlsx

Nick Phillips - 2015
phillips_nicholas@bah.com
"""

import sys
from xlrd import *
from xlwt import *
import collections
import re


def read_first_sheet(file_name):
    """
    Reads first sheet of given .xlsx file.
    :param file_name: Name of .xlsx file.
    :return: temp_sheet: text processed contents of first sheet. Row is system, column is term.
            surv_names: names of the surveillance systems
    """

    current_book = open_workbook(file_name)
    active_sheet = current_book.sheet_by_index(0)
    temp_sheet_data = []
    for rowNum in range(0, active_sheet.nrows):
        temp_row = []
        for colNum in range(2, active_sheet.ncols):
            try:
                temp_term = active_sheet.cell_value(rowNum,colNum).encode('ascii','ignore').lower().lstrip()
                if temp_term != "":
                    temp_row.append(temp_term)
            except UnicodeEncodeError as e:
                pass
        temp_sheet_data.append(temp_row)
    surv_names = active_sheet.col_values(0)

    return temp_sheet_data,surv_names

def assign_first_code(sheet_data):
    """
    Assigns "A" Level Code to each unique word.
    :param sheet_data: Sheet data extracted via read_first_sheet().
    :return: code_legend and code_sheet
    """
    pattern = re.compile('[\W_]+')
    term_list = collections.OrderedDict()
    code_sheet = []
    count = 1
    for row in sheet_data:
        temp_list = []
        for i in range(0, len(row)):
            word_parse = row[i].split()
            for word in word_parse:
                word = pattern.sub(' ',word)
                word = word.split()
                for w in word:
                    if len(w)<=1:
                        continue
                    if w not in term_list:
                        term_list[w] = "A"+str(count)
                        count += 1
                    temp_list.append(term_list[w])
        code_sheet.append(temp_list)

    return code_sheet,term_list


#-------------------------------------------------------------------------------------------

if __name__ == "__main__":
    try:
        for arg in sys.argv[1:]:
            doc_name = arg
            sheet_data,surv_names = read_first_sheet(doc_name)

            sheet_code,term_list = assign_first_code(sheet_data)

        #Create output workbook and save data.
        workbookOut = Workbook()
        sheet1 = workbookOut.add_sheet("System Codes")
        sheet2 = workbookOut.add_sheet("Code Legend")

        row = 0
        for condition in term_list:
            sheet2.write(row,0,condition)
            sheet2.write(row,1,term_list[condition])
            row+=1
        workbookOut.save(doc_name[:len(doc_name)-5]+'.out.xls')

        row_n = 0
        col = 1
        for row in sheet_code:
            sheet1.write(row_n, 0, surv_names[row_n])
            for code in row:
                sheet1.write(row_n, col, code)
                col += 1
            row_n += 1
            col = 1

        workbookOut.save(doc_name[:len(doc_name)-5]+'.out.xls')

    except SystemError as e:
        print e
        sys.exit(0)



