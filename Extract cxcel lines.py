#提取excel一整行的数据
#使用时仅需更改11、12行，会在运行程序的文件夹生成新的xls


import os
import xlwt
import xlrd
from openpyxl import load_workbook

##目的文件夹
dirpath=r'C:\Users\I\Desktop\20220625\20220708\bHLH\Eucalyptus grandis'
keyword='bHLH'

##遍历函数
def files(dirpath, suffix=['.xls', 'xlsx']):
    for root ,dirs ,files in os.walk(dirpath):
        for name in files:
            if name.split('.')[-1] in suffix:
                yield os.path.join(root, name)

if __name__ == '__main__':

    jieguo = xlwt.Workbook(encoding="ascii")  #生成excel
    wsheet = jieguo.add_sheet('sheet name') #生成sheet    
    y=0 #生成的excel的行计数
    try:
        file_list = files(dirpath)
        for filename in file_list:
            workbook = xlrd.open_workbook(filename) #读取源excel文件
            print(filename)
            sheetnum=workbook.nsheets  #获取源文件sheet数目
            for m in range(0,sheetnum):
                sheet = workbook.sheet_by_index(m) #读取源excel文件第m个sheet的内容
                nrowsnum=sheet.nrows  #获取该sheet的行数
                for i in range(0,nrowsnum):
                    date=sheet.row(i) #获取该sheet第i行的内容
                    for n in range(0,len(date)):
                        aaa=str(date[n]) #把该行第n个单元格转化为字符串，目的是下一步的关键字比对
                        print(aaa)
                        if aaa.find(keyword)>0: #进行关键字比对，包含关键字返回1，否则返回0
                            y=y+1
                            for j in range(len(date)):
                                wsheet.write(y,j,sheet.cell_value(i,j)) #该行包含关键字，则把它所有单元格依次写入入新生成的excel的第y行
        jieguo.save('jieguo.xls') #保存新生成的Excel
    except Exception as e:
        print(e)

    jieguo.save('jieguo.xls') #保存新生成的Excel        
