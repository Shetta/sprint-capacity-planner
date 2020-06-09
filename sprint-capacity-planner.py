# library modules
import datetime
import time
import openpyxl
import sys


# local include files
import credentials as cr

start_time = time.time()


def load_bank_holidays(file_dict, sheet_name):
    xlsfile = file_dict['path'] + '/' + file_dict['filename']
    print(xlsfile)
    try:
        wb = openpyxl.load_workbook(xlsfile, read_only=True, data_only=True)
    except:
        print("Unexpected error:", sys.exc_info()[0])
    else:
        load_max_col = 3
        keys = []
        all_rows = []
        sheet = wb[sheet_name]
        for firstrow in sheet.iter_rows(min_row=1, max_row=1, min_col=0, max_col=load_max_col, values_only=True):
            for key in firstrow:
                keys.append(key)
        for row in sheet.iter_rows(min_row=2, min_col=0, max_col=load_max_col, values_only=True):
            onerow = {}
            for i in range(0, load_max_col):
                onerow[firstrow[i]] = row[i]
            all_rows.append(onerow)
        print(keys)
        print(all_rows)


if __name__ == '__main__':
    now = datetime.datetime.now()
    print("Started at:", now.strftime("%Y-%m-%d %H:%M:%S"))
    print("#####----------------------------------#####")
    load_bank_holidays(cr.testexcelfile, 'Bank holidays')
    print("#####----------------------------------#####")
    print("--- Execution time: %s seconds ---" % (time.time() - start_time))

