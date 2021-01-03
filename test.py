import os
import os.path
import tkinter as tk

import xlsxwriter
import xlwt
from pptx import Presentation
from docx import Document
from win32com.client import constants, gencache

root = tk.Tk()
root.geometry('400x200')
root.title('选择检测规则')
# 第一行（两列）
row1 = tk.Frame(root)
row1.pack(fill="x")
l1 = tk.Label(row1, text='关键字信息：', width=12, height=2).pack(side=tk.LEFT)
root.name = tk.StringVar()
u1 = tk.Entry(row1, textvariable=root.name, width=20)
u1.pack()
# 第二行
row2 = tk.Frame(root)
row2.pack(fill="x")
l2 = tk.Label(row2, text='阈值：', width=12, height=2).pack(side=tk.LEFT)
root.name = tk.StringVar()
u2 = tk.Entry(row2, textvariable=root.name, width=20)
u2.pack()
# 第三行
row2 = tk.Frame(root)
row2.pack(fill="x")
l2 = tk.Label(row2, text='1、文件默认保存路径为C:\测试，请在C盘下创建此文件夹', fg='red').pack(side=tk.LEFT)
# 第四行
row2 = tk.Frame(root)
row2.pack(fill="x")
l2 = tk.Label(row2, text='2、点击“确定”后直接关闭此窗口', fg='red').pack(side=tk.LEFT)


def get():
    # 关键字信息
    data1 = u1.get()
    # 阈值
    data2 = u2.get()
    print(data1, data2)

    count = 1
    base_path = r"C:/测试/"
    txt = open(base_path + "txt.txt", "w")
    # xls = open(base_path + "1.xls", "w")
    # doc = open(base_path + "1.doc", "w")

    while count <= int(data2):
        txt.write(data1 + '\n')
        # xls.write(data1 + '\n')
        # doc.write(data1 + '\n')
        count += 1
    txt.close()
    with open(r"C:\测试\txt.txt", "r") as f:
        data = f.read()
    ####################################
    # 写入ppt/pptx
    # 设置路径
    work_path = r'C:\测试'
    os.chdir(work_path)

    # 实例化 ppt 文档对象
    prs = Presentation()

    # 选择布局
    title_slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(title_slide_layout)
    # 容器
    shapes = slide.shapes
    body_shape = shapes.placeholders[1]
    tf = body_shape.text_frame

    ### 写入正文
    new_para = tf.add_paragraph()  # 添加段落
    new_para.text = data

    # 保存 ppt和pptx
    prs.save('1.pptx')
    prs.save('1.ppt')
    #####################################
    # 写入doc/docx文件
    doc = Document()
    p = doc.add_paragraph(data)
    # p.text = data
    doc.save('1.doc')
    doc.save('1.docx')
    ######################################
    # 写入xls/xlsx文件
    txtopen = open("C:/测试/txt.txt", 'r')
    lines = txtopen.readlines()
    # 新建一个excel文件
    xls = xlwt.Workbook(encoding='utf-8', style_compression=0)
    xlsx = xlsxwriter.Workbook(base_path + '1.xlsx')

    # 新建一个sheet
    xlssheet = xls.add_sheet('sheet1')
    xlsxsheet = xlsx.add_worksheet('sheet1')
    # 写入写入a.txt，a.txt文件有N行文件
    i = 0
    for line in lines:
        xlssheet.write(i, 0, line)
        xlsxsheet.write(i, 0, line)
        i = i + 1
    xls.save(base_path + '1.xls')
    xlsx.close()

    #####################################
    # 写入pdf文件
    docx_file_path = base_path + '1.docx'
    pdf_file_path = base_path + '1.pdf'
    w = gencache.EnsureDispatch('Word.Application')
    doc = w.Documents.Open(docx_file_path, ReadOnly=1)
    doc.ExportAsFixedFormat(pdf_file_path,
                            constants.wdExportFormatPDF,
                            Item=constants.wdExportDocumentWithMarkup,
                            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)

    w.Quit(constants.wdDoNotSaveChanges)


tk.Button(root, text='确定', command=get).pack()
root.mainloop()
