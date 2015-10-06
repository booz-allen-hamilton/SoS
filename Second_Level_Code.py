"""
Script to process First_Level_Code.py output and generate second level codes.
Reads output from *.*Column*.out.xls file and processes into intermediate *.*column*.matrix.xls
Performs PCA on the matrix information (via SVD).
Clusters with K-means.
Assigns B Codes to K-Means output.
"""

import sys
import xlrd
import numpy as np


def read_code_file(file_name):
    """
    Reads code file and converts into matrix.
    0 for lack of code, 1 for code.
    :param file_name:
    :return: sys_names,raw_matrix
    """
    current_book = xlrd.open_workbook(file_name)
    active_sheet = current_book.sheet_by_index(0)
    code_sheet = current_book.sheet_by_index(0)
    sys_names = []

    #Create matrix of zeros for raw data
    raw_matrix = np.zeros((active_sheet.nrows,code_sheet.nrows))

    #Iterate over values of sys name and "A" Codes
    for i in range(0,active_sheet.nrows):
        row = active_sheet.row_values(i)
        #Get sys name
        sys_names.append(row[0])
        row = row[1:]
        row = [int(num[1:]) for num in row if num]
        #Set raw matrix value
        for j in row:
            raw_matrix[i,j-1] = 1

    return sys_names,raw_matrix


if __name__ == "__main__":
    code_file_name = "SoS_Master_Copy.conditions.out.xls" #sys.argv[1]
    sys_names,raw_matrix = read_code_file(code_file_name)

    print raw_matrix