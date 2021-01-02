# 姓名：陈安岳，职称：访问教授，住址：慧园3栋519，电话：0755-88018688


# d = {'姓名':'陈安岳','职称':'访问教授','住址':'慧园3栋519','电话':'0755-88018688'}

list ='姓名：陈安岳，职称：访问教授，住址：慧园3栋519，手机号：15147176954'
count = 1
data = open(r"E:\测试\手机号.xls","w")
while count<6001:
    # print(list)
    # data.write(list)
    print(list, file=data)
    count += 1
data.close()
