import os
import tkinter as tk
from pptx import Presentation
from pptx.util import Cm, Pt

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


def get():
    # 关键字信息
    data1 = u1.get()
    # 阈值
    data2 = u2.get()
    print(data1,data2)

    count = 1
    base_path = "E:/测试/"
    address = open(base_path + "1.xls", "w", encoding="utf-8")
    while count <= int(data2):
        address.write(data1 + '\n')
        count += 1
    address.close()

tk.Button(root, text='确定', command=get).pack()
root.mainloop()

