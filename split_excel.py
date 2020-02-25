import xlrd
import xlwt
import time

excel_add = input('请粘贴待拆分表的路径（包含文件名）：')
save_add = input('请粘贴存取路径：')
xlsx = xlrd.open_workbook(str(excel_add))
table = xlsx.sheet_by_index(0)


for i in range(1,table.col_values(0).__len__()):
    xlsx2 = xlwt.Workbook()
    sheetq = xlsx2.add_sheet('test1')
    for m in range(table.row_values(0).__len__()):
        sheetq.write(0, m, table.cell_value(0, m))
        sheetq.write(1, m, table.cell_value(i,m))
        xlsx2.save('{}/{}拆分表{}.xls'.format(save_add,table.cell_value(i,0),time.strftime("%Y-%m", time.localtime())))