import os
import xlrd
import xlwt

GA_S = 256
MID_S = 59.952
SJM_S = 193
WY_S = 145
SC_S = 243
MGM_S = 59.953


def generateCSV(input_path, output_path):
    files = os.listdir(input_path)
    newWorkBook = xlwt.Workbook()
    sheet = newWorkBook.add_sheet('main')
    k = 0
    divisor = -1
    for file in files:
        data_excel = xlrd.open_workbook(input_path + '/' + file)
        data_excel = data_excel.sheets()[0]
        n_rows = data_excel.nrows
        for i in range(n_rows):
            GA_P = data_excel.cell_value(rowx=i, colx=1)
            GA_V = data_excel.cell_value(rowx=i, colx=2)
            MID_P = data_excel.cell_value(rowx=i, colx=4)
            MID_V = data_excel.cell_value(rowx=i, colx=5)
            SJM_P = data_excel.cell_value(rowx=i, colx=7)
            SJM_V = data_excel.cell_value(rowx=i, colx=8)
            WY_P = data_excel.cell_value(rowx=i, colx=10)
            WY_V = data_excel.cell_value(rowx=i, colx=11)
            SC_P = data_excel.cell_value(rowx=i, colx=13)
            SC_V = data_excel.cell_value(rowx=i, colx=14)
            MGM_P = data_excel.cell_value(rowx=i, colx=16)
            MGM_V = data_excel.cell_value(rowx=i, colx=17)
            tot_V = GA_V + MID_V + SJM_V + WY_V + SC_V + MGM_V
            if tot_V == 0:
                mei = 0
            else:
                mei = (GA_P*GA_V*GA_S + MID_P*MID_V*MID_S + SJM_P*SJM_V*SJM_S + WY_P*WY_V*WY_S + SC_P*SC_V*SC_S + MGM_P*MGM_V*MGM_S) / tot_V
            if divisor < 0:
                divisor = mei
            mei = mei / divisor
            sheet.write(k, 0, data_excel.cell_value(rowx=i, colx=0))
            sheet.write(k, 1, mei*1000)
            k += 1


    newWorkBook.save(output_path + '/' + 'MEI' + '.xls')

    print("Finished, well done!")
    return

input_path = 'D:/Desktop/CISC7014 PROJECT/MEI_PROJECT/Generate_Date'
output_path = 'D:/Desktop/CISC7014 PROJECT/MEI_PROJECT/Generate_Mei'

generateCSV(input_path, output_path)

# input_path = 'D:/Desktop/CISC7014 PROJECT/MEI_PROJECT/GENDATE'
# output_path = 'D:/Desktop/CISC7014 PROJECT/MEI_PROJECT/GENTEMP'
#
# generateCSV(input_path, output_path)
