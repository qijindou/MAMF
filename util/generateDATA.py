import os
import xlrd
import xlwt

def generateCSVdate(input_path, output_path):
    files = os.listdir(input_path)
    for file in files:
        newWorkBook = xlwt.Workbook()
        sheet = newWorkBook.add_sheet('main'+file)
        path_up = input_path + '/' + file + '/'
        equityPath = []
        equityPath.append(path_up + 'Equities_27_' + file + '.xlsx')
        equityPath.append(path_up + 'Equities_200_' + file + '.xlsx')
        equityPath.append(path_up + 'Equities_880_' + file + '.xlsx')
        equityPath.append(path_up + 'Equities_1128_' + file + '.xlsx')
        equityPath.append(path_up + 'Equities_1928_' + file + '.xlsx')
        equityPath.append(path_up + 'Equities_2282_' + file + '.xlsx')
        k = 0
        n_prev_row = 0
        for path in equityPath:
            data_excel = xlrd.open_workbook(path)
            data_excel = data_excel.sheets()[0]
            n_rows = data_excel.nrows
            n_cols = data_excel.ncols
            if n_rows != n_prev_row and n_prev_row != 0:
                print("wrong:" + path)
                print(n_rows)
            for i in range(1, n_rows):
                for j in range(n_cols):
                    if j != 0:
                        sheet.write(i-1, j+k, float(data_excel.cell_value(rowx=i, colx=j)))
                    else:
                        sheet.write(i-1, j+k, data_excel.cell_value(rowx=i, colx=j))
            k += n_cols
            n_prev_row = n_rows

        newWorkBook.save(output_path + '/' + file + '.xls')

    print("Finished, well done!")
    return


def generateCSVstock(input_path, output_path, equityName):
    files = os.listdir(input_path)
    newWorkBook = xlwt.Workbook()
    sheet = newWorkBook.add_sheet('main')
    k = 0
    for file in files:
        path= input_path + '/' + file + '/' + equityName + '_' + file + '.xlsx'
        data_excel = xlrd.open_workbook(path)
        data_excel = data_excel.sheets()[0]
        n_rows = data_excel.nrows
        n_cols = data_excel.ncols
        for i in range(1, n_rows):
            for j in range(n_cols):
                if j != 0:
                    sheet.write(k, j, float(data_excel.cell_value(rowx=i, colx=j)))
                else:
                    sheet.write(k, j, data_excel.cell_value(rowx=i, colx=j))

            k += 1

    newWorkBook.save(output_path + '/' + equityName + '.xls')

    print("Finished, well done!")
    return


input_path_date = 'D:/Desktop/CISC7014 PROJECT/MEI_PROJECT/DATA'
output_path_date = 'D:/Desktop/CISC7014 PROJECT/MEI_PROJECT/Generate_Date'
input_path_stock = input_path_date
output_path_stock = 'D:/Desktop/CISC7014 PROJECT/MEI_PROJECT/Generate_Stock'

generateCSVdate(input_path_date, output_path_date)
generateCSVstock(input_path_stock, output_path_stock, 'Equities_27')
generateCSVstock(input_path_stock, output_path_stock, 'Equities_200')
generateCSVstock(input_path_stock, output_path_stock, 'Equities_880')
generateCSVstock(input_path_stock, output_path_stock, 'Equities_1128')
generateCSVstock(input_path_stock, output_path_stock, 'Equities_1928')
generateCSVstock(input_path_stock, output_path_stock, 'Equities_2282')

# input_path_date = 'D:/Desktop/CISC7014 PROJECT/MEI_PROJECT/TEMP'
# output_path_date = 'D:/Desktop/CISC7014 PROJECT/MEI_PROJECT/GENDATE'
#
# generateCSVdate(input_path_date, output_path_date)
# input_path_stock = 'D:/Desktop/CISC7014 PROJECT/MEI_PROJECT/TEMP'
# output_path_stock = 'D:/Desktop/CISC7014 PROJECT/MEI_PROJECT/GENTEMP'
#
# generateCSVstock(input_path_stock, output_path_stock, 'Equities_27')
# generateCSVstock(input_path_stock, output_path_stock, 'Equities_200')
# generateCSVstock(input_path_stock, output_path_stock, 'Equities_880')
# generateCSVstock(input_path_stock, output_path_stock, 'Equities_1128')
# generateCSVstock(input_path_stock, output_path_stock, 'Equities_1928')
# generateCSVstock(input_path_stock, output_path_stock, 'Equities_2282')
