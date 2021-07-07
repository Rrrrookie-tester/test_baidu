"""
@Project: AutoFrameWork_WebTps   
@Description: TODO          
@Time:2021/7/7 9:19       
@Author:zexin                
 
"""
from time import sleep

import xlwt as xlwt
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException


def search():
    flag = 0  # 标志位。用于判断是否有搜索结果
    res_list = []  # 定义结果集
    driver = webdriver.Chrome(r'D:\chromedriver\chromedriver_win32\chromedriver.exe')  # 初始化webdriver
    driver.get('http://www.baidu.com')  # 打开百度页面
    try:
        search_input = driver.find_element_by_id('kw')  # 获取到页面搜索框元素
    except NoSuchElementException:
        print("can not find search input !")  # 捕获异常并打印提示
    try:
        search_button = driver.find_element_by_id('su')  # 获取到页面搜索按钮元素
    except NoSuchElementException:
        print("can not find search button !")
    try:
        search_input.send_keys('可转债')  # 向输入框传值“可转债”
        search_button.click()  # 点击搜索按钮
    except Exception as e:
        print(e)
    sleep(5)
    # 搜索步骤完成，开始验证搜索结果
    search_title = driver.title  # 获取当前页面标题
    if '可转债' in search_title:
        print('搜索结果存在')
        flag = 1                 # 判断“可转债”是否在标题中，若在则存在搜索结果
    else:
        print('搜索结果不存在')
        quit()
    if flag == 1:               # 当存在搜索结果时，进行后续操作
        first_page_res = driver.find_elements_by_css_selector(".c-container > h3 > a")  # 获取到第一页的10条搜索结果
        for res in first_page_res:
            href = res.get_attribute('href')
            res_list.append(href)    # 将第一页结果添加到结果集中
        try:
            next_page_button = driver.find_element_by_class_name('n')  # 定位跳转到第二页的按钮
        except NoSuchElementException:
            print("can not find next page button !")
        try:
            next_page_button.click()  # 点击跳转到第二页
        except:
            print("click next page button error !")
        sleep(5)  # 强制等待5s
        second_page_res = driver.find_elements_by_css_selector(".c-container > h3 > a")  # 获取到第二页的10条搜索结果
        for res in second_page_res:
            href = res.get_attribute('href')
            res_list.append(href)  # 将第二页结果添加到结果集中
    driver.quit()   # 退出、关闭浏览器
    return res_list # 返回结果集


def write_result(file_path, datas):   # 将结果集写入excel
    f = xlwt.Workbook(encoding='utf-8')
    sheet1 = f.add_sheet(u'Top20', cell_overwrite_ok=True)  # 创建sheet
    i = 0
    for data in datas:
        sheet1.write(i, 0, data)
        i = i + 1
    f.save(file_path)


result = search()
write_result(file_path=r"xxxxxxxx", datas=result)