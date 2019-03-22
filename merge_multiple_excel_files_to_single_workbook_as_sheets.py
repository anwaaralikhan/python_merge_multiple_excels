import xlsxwriter
import os
import csv


#################################################################
#################################################################
directory_location ='C:\folder_with_multiple_excel_files'
merge_file_location ='C:\output\merged_excel_workbook.xlsx'


#################################################################
#################################################################
#load all files
files_in_directory = os.listdir(directory_location)


#################################################################
#################################################################
# create merge workbook
workbook = xlsxwriter.Workbook(merge_file_location)


#################################################################
#################################################################
for file in files_in_directory:
    fName = directory_location  + '\\' + file
    basic_name = file.split('_')  # this part is optional, my files were _ separated with time_stamp ex: file1_03222019
    sheet_name = basic_name[0]
    worksheet = workbook.add_worksheet(sheet_name)
    with open(fName, encoding='utf-8') as csv_file:
        csv_reader = csv.reader(csv_file)
        row_count = 0
        for row in csv_reader:
            col_count = 0
            for col in row:
                worksheet.write(row_count,col_count,col)
                col_count +=1
            row_count +=1

            
#################################################################
#################################################################

workbook.close()
