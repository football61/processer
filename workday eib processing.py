import openpyxl
from openpyxl.utils import get_column_letter
import datetime
import pandas as pd

def update_excel_file(file_path, target_cell11, target_cell12, target_cell21, target_cell22, df):
    workbook = openpyxl.load_workbook(file_path)
    sheet1 = workbook["Edit License"]
    sheet2 = workbook["Hire Employee"]

    column11 = target_cell11[0]
    print(column11)
    start_row11 = int(target_cell11[1:])
    column12 = target_cell12[0]
    start_row12 = int(target_cell12[1:])




    column21 = target_cell21[0]
    start_row21 = int(target_cell21[1:])
    column22 = target_cell22[0]
    start_row22 = int(target_cell22[1:])    

    for index, value in enumerate(df['Spreadsheet Key*']):
        cell11 = sheet1["{}{}".format(column11, start_row11 + index)]
        cell11.value = value
        
    for index, value in enumerate(df['Applicant*']):
        cell12 = sheet1["{}{}".format(column12, start_row12 + index)]
        cell12.value = value        
        
        
        

    for index, value in enumerate(df['Universal ID']):
        cell21 = sheet2["{}{}".format(column21, start_row21 + index)]
        cell21.value = value




# "concatenates" two columns
    for index, value in enumerate(df.apply(lambda row: str(row['Universal ID']) + str(row['Applicant*']), axis=1)):
        cell22 = sheet2["{}{}".format(column22, start_row22 + index)]
        cell22.value = value              


    workbook.save('outputeib ' + datetime.datetime.now().strftime("%Y_%m_%d %H%M") + '.xlsx')


file_path = "eibtemplate.xlsx"
df = pd.read_excel('input file.xlsx')
print(df)
'''
   Spreadsheet Key*  Applicant*  Universal ID  Country*
0          724572457    78612542            17  usa
1           24524572    75412457            17  usa
2           34567563    24572457            18  canada
3          245675368    75245788            19  canada
4            5638356    24567245            20  mexico
5          354653688    24578458            21  mexico
'''


target_cell11 = "b6"
target_cell12 = "c6"


target_cell21 = 'b6'
target_cell22 = 'g6'


update_excel_file(file_path, target_cell11, target_cell12, target_cell21, target_cell22, df)

