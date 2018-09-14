from selenium import webdriver
import time
import requests
import xlwt
from lxml import etree

# 创建工作簿
wbk = xlwt.Workbook(encoding='utf-8', style_compression=0)
# 创建工作表
sheet = wbk.add_sheet('sheet 1', cell_overwrite_ok=True)
# 创建表头
table_top_list = [u"业务主题", "对应客户", "类型", "状态", "创建日期", "沟通备注"]
# 写表头数据
for c, top in enumerate(table_top_list):
    sheet.write(0, c, top)

driver = webdriver.Chrome()
url = 'http://www.400cx.com/platform/login.jsp'
driver.get(url)

'''窗口最大化'''
# driver.maximize_window()
driver.find_element_by_xpath('.//input[@id="userNameIpt"]').send_keys('2014@4006875598')
# driver.find_element_by_id('userNameIpt')
driver.find_element_by_xpath('.//input[@id="password"]').send_keys('hp@@123456')
pw = int(input("请输入验证码:"))
driver.find_element_by_xpath('.//input[@id="pwdInput"]').send_keys(pw)
driver.find_element_by_id('btnSubmit').click()
time.sleep(1)
# driver.switch_to.frame('customerCenterframe')
driver.find_element_by_xpath('.//a[@url="/platform/saleRecord/saleRecordList.do?recordType=1"]').click()
time.sleep(1)
# iframe  01
driver.switch_to.frame('customerCenterframe')


#  只能访问第一条数据，考虑取到所有的a标签，遍历点击事件
#  driver.find_element_by_xpath('//table[@class="clist_01"]//tr//a').click()


def save_msg(i):
    time.sleep(1)
    #  跳到原始页面 02
    driver.switch_to.default_content()
    try:
        driver.find_element_by_xpath('//div[@class="sale_conLtit"]/span[2]')
    except:
        cus_msg = ""
    else:
        cus_msg = driver.find_element_by_xpath('//div[@class="sale_conLtit"]/span[2]').text
    try:
        driver.find_elements_by_class_name('fw_b')[-1]
    except:
        topic_msg = ""
    else:
        topic_msg = driver.find_element_by_class_name('fw_b')[-1].text
    try:
        texts = driver.find_elements_by_class_name('td_right')
    except:
        print('no data in table')
    else:
        a = texts[0].text if texts else ""
        b = texts[1].text if texts else ""
        date = texts[3].text if texts else ""
        d = texts[2].text if texts else ""
        sheet.write(i, 0, topic_msg)
        sheet.write(i, 1, cus_msg)
        sheet.write(i, 2, a)
        sheet.write(i, 3, b)
        sheet.write(i, 4, date)
        sheet.write(i, 5, d)
        wbk.save(r'C:\Users\Administrator\Desktop\400msg.xls')
        print(topic_msg, cus_msg, i, date)
		#  点击关闭按钮，关闭详细页
        driver.find_element_by_class_name('tabs-close').click()
        time.sleep(1)
        driver.switch_to.frame('customerCenterframe')


def login_detail_page(i):
    for j in range(i, i + 10):
        #  点击进入第二层页面
        driver.find_element_by_xpath('//table[@class="clist_01"]//tr[{}]//a'.format(j % 10 + 1)).click()
        save_msg(j)
        print('-+' * 50)
    return i + 10


# # 第一页
# i = login_detail_page(0)
# driver.find_element_by_class_name('next').click()
#
# # 第二页
# i = login_detail_page(i)
# driver.find_element_by_class_name('next').click()
#
# # 第三页
# i = login_detail_page(i)
# driver.find_element_by_class_name('next').click()
#
# # 第四页
# i = login_detail_page(i)
# driver.find_element_by_class_name('next').click()
#
# # 第五页
# i = login_detail_page(i)


# 此处递归调用 打开能取到所有的数据
def next_page_fun(i):
    iter_time = login_detail_page(i)
    driver.find_element_by_class_name('next').click()
    return next_page_fun(iter_time)


next_page_fun(0)
