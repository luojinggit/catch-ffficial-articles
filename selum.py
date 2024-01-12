import calendar
import json
import os
import re
from time import sleep

import xlwt
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By

# 获取当前日期的元组表示
current_date = calendar.datetime.datetime.now()
# 获取当前年份
year = current_date.year
# 获取当前月份
month = current_date.month
# 获取当前日
day = current_date.day

exDate = str(year) + '-' + str(month) + '-' + str(day)

official_accounts = input("请输入需要爬取的公众号？不输入直接回车默认为‘信通院’,如需爬取多个则用逗号隔开")
if official_accounts == '':
    official_accounts = '信通院'

answer = input("是否只爬取当天的公众号文章？(y/n)")
if answer.lower() != "y":
    s1 = input("请输入截止日期？（示例：2023-01-01）")
    date_regex = r'^\d{4}-\d{2}-\d{2}$'
    while not re.match(date_regex, s1):
        s1 = input("日期格式错误，请重新输入：")
    exDate = s1

# 添加保持登录的数据路径：安装目录一般在C:\Users\****\AppData\Local\Google\Chrome\User Data
user_data_dir = r'C:\Users\luojing\AppData\Local\Google\Chrome\User'
# 这是一个选项类
user_option = webdriver.ChromeOptions()
# 添加浏览器用户数据
user_option.add_argument(f'--user-data-dir={user_data_dir}')
# 实例化浏览器（带上用户数据）
driver = webdriver.Chrome(options=user_option)

driver.get("https://mp.weixin.qq.com/")
driver.set_window_size(1500, 1200)
sleep(30)

button0 = driver.find_element(By.XPATH, '//*[@id="js_index_menu"]/ul/li[3]/span')
button0.click()

sleep(2)

# 定位草稿箱按钮
button = driver.find_element(By.XPATH, '//*[@id="js_level2_title"]/li[1]/ul/li[1]/a/span/span')
# 执行单击操作
button.click()
sleep(2)

# 定位新的创作
collect = driver.find_element(By.XPATH, '//*[@id="js_main"]/div[3]/div[2]/div/div/div/div[2]/div/div/div[1]')
# 悬停至收藏标签处
ActionChains(driver).move_to_element(collect).perform()
sleep(2)

# 定位新建公众号
button = driver.find_element(By.XPATH, '//*[@id="js_main"]/div[3]/div[2]/div/div/div/div[2]/div/div/div[2]/ul/li[1]')
button.click()

# 切换窗口
driver.switch_to.window(driver.window_handles[-1])
sleep(2)

# 定位超链接按钮
button = driver.find_element(By.XPATH, '//*[@id="js_editor_insertlink"]')
# 执行单击操作
button.click()
sleep(2)

accounts = official_accounts.split(",")
datas_list = []
for i in range(len(accounts)):
    account_title = accounts[i]
    # 定位选择其他公众号按钮，也就是想获取得公众号名称
    button = driver.find_element(By.XPATH,
                                 '//*[@id="vue_app"]/div[2]/div[1]/div/div[2]/div[2]/form[1]/div[4]/div/div/p/div/button')
    # 执行单击操作
    button.click()
    sleep(2)

    # 定位搜索输入框
    text_label = driver.find_element(By.XPATH,
                                     '//*[@id="vue_app"]/div[2]/div[1]/div/div[2]/div[2]/form[1]/div[4]/div/div/div/div/div[1]/span/input')
    text_label.clear()
    text_label.send_keys(account_title)
    sleep(2)

    # 定位搜索按钮
    button = driver.find_element(By.XPATH,
                                 '//*[@id="vue_app"]/div[2]/div[1]/div/div[2]/div[2]/form[1]/div[4]/div/div/div/div/div[1]/span/span/button[2]')
    # 执行单击操作
    button.click()
    sleep(2)

    # 定位第一条搜索结果
    button = driver.find_element(By.XPATH,
                                 '//*[@id="vue_app"]/div[2]/div[1]/div/div[2]/div[2]/form[1]/div[4]/div/div/div/div[2]/ul/li[1]/div[1]')
    # 执行单击操作
    button.click()
    sleep(5)

    page_number_list = driver.find_elements(By.CLASS_NAME, 'weui-desktop-pagination__num')
    page_number = page_number_list[1].text
    for i in range(int(page_number)):
        # 定位数据
        datas = driver.find_elements(By.CLASS_NAME, 'inner_link_article_item')
        is_over = False
        # 向列表中添加数据
        for i in datas:
            row_list = []
            row_list.append(account_title)
            title = i.find_element(By.CLASS_NAME, 'inner_link_article_title').text
            row_list.append(title)
            date = i.find_element(By.CLASS_NAME, 'inner_link_article_date').text
            row_list.append(date)
            link = i.find_element(By.TAG_NAME, 'a').get_attribute('href')
            row_list.append(link)
            if exDate > date:
                is_over = True
                break
            datas_list.append(row_list)

        # 实现翻页
        button = driver.find_elements(By.XPATH,
                                      '//*[@id="vue_app"]/div[2]/div[1]/div/div[2]/div[2]/form[1]/div[5]/div/div/div[3]/span[1]/a')[
            -1]
        button.click()
        sleep(4)
        if is_over:
            break

book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('data', cell_overwrite_ok=True)
col = ('公众号', '标题', '发表时间', '文章链接')
for i in range(0, 3):
    sheet.write(0, i, col[i])

json_array = []
for i in range(len(datas_list)):
    data = datas_list[i]
    for j in range(0, 4):
        sheet.write(i + 1, j, data[j])
    json_obj = {
        'account': data[0],
        "title": data[1],
        "time": data[2],
        "url": data[3]
    }
    json_array.append(json_obj)
sheet.write(len(datas_list) + 2, 0, json.dumps(json_array, ensure_ascii=False))

savePath = os.path.join("D:", exDate + ".xls")
book.save(savePath)

driver.quit()

print("执行完成，窗口将在5秒内关闭")
for i in range(5, 0, -1):
    print(i)
