import xlrd
import xlwt
import os
from os.path import abspath
import sys
import pymongo


full_url = "mongodb://127.0.0.1:27017/test"
conn = pymongo.MongoClient(full_url)['excel_import']


def main(file_path):
    f_path = abspath(file_path)
    # Open the workbook
    xl_workbook = xlrd.open_workbook(f_path)

    # List sheet names, and pull a sheet by name
    sheet_names = xl_workbook.sheet_names()

    # for each sheet
    for sheet in sheet_names:
        # individual sheet object
        xl_sheet = xl_workbook.sheet_by_name(sheet)

         # Number of columns
        num_cols = xl_sheet.ncols

        # Iterate through rows
        for row_idx in range(0, xl_sheet.nrows):
            # this is a header
            if row_idx == 0:
                header = get_row_values(xl_sheet, num_cols, row_idx)

            # these are data values
            else:
                row = get_row_values(xl_sheet, num_cols, row_idx)
                # build the dictionary for mongo
                d = {}
                for i, val in enumerate(row):
                    d[header[i]] = val

                # insert into mongo
                print(f'inserting row: {row_idx} in sheet {sheet}')
                conn[sheet].insert_one(d)



def get_row_values(xl_sheet, num_cols, row_idx):
    r = []
    # Iterate through columns
    for col_idx in range(0, num_cols):
        # Get cell object by row, col
        cell_obj = xl_sheet.cell(row_idx, col_idx)
        # push
        r.append(cell_obj.value)

    return r


#----------------------------------------------------------------------
if __name__ == "__main__":
    main("demo.xlsx")
