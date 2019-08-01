# Excel Sheet CSV decomposer
import os
import sys
import subprocess
import csv
print("\n<Excel_Sheet_to_CSV> by.Junsu Shin @https://github.com/ashtapor25\n")

try:
    import xlrd
except ImportError:
    print("Installing Required Packages...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'xlrd'])
    import xlrd

def excel_to_csv():
    # take the file name as input
    target_file = input("Write the file name to break into csvs (ex. myfile.xlsx, myfile.xls)\n")

    # check if the file exists
    exists = os.path.isfile(target_file)
    if not exists:
        print(target_file+" does not exist in the directory")
        return

    no_filename_chars = ['<', '>', '?', '[', ']', ':', '|', '*']
    transtable = str.maketrans('','','<>?[]:|*')

    # create a folder in the home directory for the CSVs
    folder_name = target_file.replace(".xlsx", "").replace(".xls", "")
    os.mkdir(folder_name)
    print("folder "+folder_name+" created in directory")

    print('characters <>?[]:|* will be removed from sheet names when making csv files')
    # create a CSV file for every sheet
    with xlrd.open_workbook(target_file) as wb:
        sheet_lst = wb.sheets()
        for sheet in sheet_lst:
            sheetname = sheet.name.translate(transtable)
            with open(folder_name+"/"+sheetname+'.csv', 'w', newline="") as f:
                c = csv.writer(f)
                for r in range(sheet.nrows):
                    c.writerow(sheet.row_values(r))
            print(sheetname+'.csv created')

    return

excel_to_csv()
input("Press Enter to Exit Program")