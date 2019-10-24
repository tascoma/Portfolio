import openpyxl

wb = openpyxl.load_workbook('Portfolio Management.xlsx')
sheetACP = wb['Adj Close Prices']
sheetReturns = wb['Returns']
columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']


for col in columns:
    x = 6
    cell1 = sheetACP[col + str(x)].value
    while cell1:
        cell2 = sheetACP[col + str(x - 1)].value 
        sheetReturns[col + str(x - 1)] = cell1/cell2 - 1
        x += 1
        cell1 = sheetACP[col + str(x)].value


wb.save('Portfolio Management.xlsx')

