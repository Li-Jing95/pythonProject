import os
import webbrowser
import zipfile
from selenium import webdriver
import time


z = zipfile.ZipFile(r'C:\Users\Administrator\Desktop/1234.zip','r')
# 返回所有文件夹和文件
# zip_list = z.namelist()

# for zip in zip_list:
#     # print(zip_list)
#     # print(zip.encode('utf-8'))
#     try:
#         zip = zip.encode('cp437').decode('gbk')
#     except:
#         zip = zip.encode('utf-8').decode('utf-8')
#     print(zip)
z.extractall(r"C:\Users\Administrator\Desktop")
z.close()


driver = webdriver.Chrome()#打开浏览器
driver.get('https://mail.qq.com/')#打开百度官网
time.sleep(5)

driver.switch_to.frame('login_frame')
# user_name = driver.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/form/div[1]/div/input')
driver.find_element_by_id('u').clear()
driver.find_element_by_id('u').send_keys('24663773')
time.sleep(2)
driver.find_element_by_xpath('/html/body/div[1]/div[4]/div/div/div[2]/form/div[2]/div[1]/input').send_keys('13223326073,lj')
time.sleep(2)

driver.find_element_by_id('login_button').click()


