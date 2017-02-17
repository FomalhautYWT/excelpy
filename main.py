import xlrd
import os
import xlwt
def makexlsx(filename,rowvalue):
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    borders.bottom_colour = 0x3A
    borders.top_colour = 0x3A
    borders.left_colour = 0x3A
    borders.right_colour = 0x3A




    style = xlwt.XFStyle()
    alignment = xlwt.Alignment()
    alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT
    alignment.vert = xlwt.Alignment.VERT_CENTER

    style.alignment = alignment
    style.borders = borders

    rdfile = xlrd.open_workbook('D:\\xxx\\hhh.xlsx')
    table = rdfile.sheet_by_index(0)
    nrows = table.nrows
    file = xlwt.Workbook()
    sheet = file.add_sheet('sheet1', cell_overwrite_ok=True)
    for i in range(3):
        sheet.col(i).width = 256*30
    sheet.col(3).width = 256*30
    sheet.col(4).width = 256*20
    for i in range(23):
        sheet.row(i).height = 100
    for i in range(nrows):
        rw = table.row_values(i)
        if (i != 17) and (i!= 18) and (i!= 19):
            for j in range(3):
                sheet.write(i, j, str(rw[j]), style)
        else :
            for j in range(5):
                sheet.write(i,j,str(rw[j]),style)
    sheet.write(18,4,'   年   月   日',style)
    sheet.write(19, 4, '   年   月   日', style)
    for i in range(20,23):
        for j in range(3,5):
            sheet.write(i,j,' ',style)
    sheet.write(0,1,rowvalue[4],style)

    sheet.row(1).set_cell_text(1,str(rowvalue[2]).replace('.0',''),style)

    sheet.write(2, 1, str(rowvalue[6]) + ',' + str(rowvalue[7]),style)
    sheet.write(4,1,rowvalue[12],style)
    sheet.write(5, 1, rowvalue[11],style)
    sheet.write(6,1,(str(rowvalue[13]) + ',' + str(rowvalue[14])),style)
    sheet.write_merge(0,0,3,4, rowvalue[15],style)
    sheet.write_merge(1,1,3,4, str(rowvalue[4]) + '-预覆盖工单',style)
    sheet.write_merge(2,2,3,4,rowvalue[5],style)
    sheet.write_merge(3,3,3,4,'现场反馈数据',style)
    for i in range(4,17):
        sheet.write_merge(i,i,3,4,' ',style)
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT
    alignment.vert = xlwt.Alignment.VERT_CENTER
    font = xlwt.Font()
    font.bold = 2
    font.height = 270
    style.alignment = alignment
    style.font = font
    sheet.write(3,0,'数据类型',style)
    sheet.write(3,1,'综资现有数据',style)
    sheet.write(3,2,'初步判断数据',style)
    sheet.write_merge(3,3,3,4,'现场反馈数据',style)

    file.save(filename)

def main():
    xlsfile = 'D:\\xxx\\1.xlsx'
    source_excel = xlrd.open_workbook(xlsfile)
    table = source_excel.sheet_by_index(0)
    nrows = table.nrows
    for i in range(nrows):
        print(table.row_values(i))
        rowvalue = table.row_values(i)
        path = 'D:\\xxx\\baobiao\\' + rowvalue[1]
        if (not os.path.exists(path)):
            os.mkdir(path)
        flag = False
        for j in range(nrows):
            if (j != i) and (table.row_values(i)[11] == table.row_values(j)[11] and table.row_values(i)[1] == table.row_values(j)[1]):
                flag = True
                break
        print (str(table.row_values(i)))
        rw11 = str(table.row_values(i)[11]).replace('/','')
        if flag :
            path = path + '\\' + rowvalue[1] + rw11
            if not os.path.exists(path):
                os.mkdir(path)
        if flag :
            filename = 'D:/xxx/baobiao/'
            filename = filename + rowvalue[1]
            filename = filename + '/' + rowvalue[1] + rw11
            filename = filename + '/' + rowvalue[1]
            it =  str(rowvalue[2]).replace('.0','')
            filename = filename + it + '.xls'
        else :
            filename = 'D:/xxx/baobiao/' + rowvalue[1] + '/' + rowvalue[1] + str(rowvalue[2]).replace('.0','') + '.xls'


        makexlsx(filename,rowvalue)


if __name__ == "__main__":
    main()