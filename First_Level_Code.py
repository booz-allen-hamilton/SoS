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
import csv
from nltk.corpus import stopwords


def read_first_sheet(master_file_name,file_name):
    """
    Reads first sheet of given .xlsx file.
    :param file_name: Name of .xlsx file.
    :return: temp_sheet: text processed contents of first sheet. Row is system, column is term.
            surv_names: names of the surveillance systems
    """
    temp_sheet_data = []
    with open(file_name,'rb') as csvfile:
        sheet_reader = csv.reader(csvfile)
        for row in sheet_reader:
            temp_sheet_data.append(row)

    master_book = open_workbook(master_file_name)
    active_sheet = master_book.sheet_by_index(0)
    surv_names = active_sheet.col_values(0)

    return temp_sheet_data,surv_names

def assign_first_code(sheet_data):
    """
    Assigns "A" Level Code to each unique word.
    :param sheet_data: Sheet data extracted via read_first_sheet().
    :return: code_legend and code_sheet
    """

    stop = stopwords.words('english')
    pattern = re.compile('[\W_]+')
    term_list = collections.OrderedDict()
    code_sheet = []
    count = 1
    for row in sheet_data:
        temp_list = []
        for word in row:
            word = pattern.sub(' ',word).strip()
            if len(word)<=2:
                continue
            if word in stop:
                continue
            if word not in term_list:
                term_list[word] = "A"+str(count)
                count += 1
            temp_list.append(term_list[word])
        code_sheet.append(temp_list)

    return code_sheet,term_list


#-------------------------------------------------------------------------------------------

if __name__ == "__main__":
    try:
        master_name = sys.argv[1]
        for arg in sys.argv[2:]:
            doc_name = arg
            sheet_data,surv_names = read_first_sheet(master_name,doc_name)

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

            row_n = 0
            col = 1
            for row in sheet_code:
                sheet1.write(row_n, 0, surv_names[row_n])
                for code in row:
                    sheet1.write(row_n, col, code)
                    col += 1
                row_n += 1
                col = 1

            workbookOut.save(doc_name[:len(doc_name)-4]+'.out.xls')

    except SystemError as e:
        print e
        sys.exit(0)



