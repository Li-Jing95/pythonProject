import xlwt #写入文件

fopen=open("E:/测试/txt.txt",'r')
lines=fopen.readlines()
#新建一个excel文件
file=xlwt.Workbook(encoding='utf-8',style_compression=0)
#新建一个sheet
sheet=file.add_sheet('data')

#写入写入a.txt，a.txt文件有20000行文件
i=0
for line in lines:
    sheet.write(i,0,line)
    i=i+1

file.save('minni.xls')