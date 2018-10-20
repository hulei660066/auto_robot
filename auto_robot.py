# -*- coding: utf-8 -*
import time

import xlrd
import xlwt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


def login():
    global driver
    driver = webdriver.Chrome()
    driver.get("https://jxc.lixiaocrm.com/dashboard")
    locator = (By.ID, "dashboard_index")
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable(locator))
    driver.get("https://jxc.lixiaocrm.com/products")


def write_output_excel(results):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('xmq_love')
    row = 0
    for result in results:
        col = 0
        for val in result:
            ws.write(row, col, val)
            col += 1
        row += 1
    wb.save('output.xls')


def test():
    # workbook = xlsxwriter.Workbook('Expenses01.xlsx')
    # worksheet = workbook.add_worksheet()
    # for item, cost in (expenses):
    #     worksheet.write(row, col, item)
    #     worksheet.write(row, col + 1, cost)
    #     row += 1
    #
    # # Write a total using a formula.
    # worksheet.write(row, 0, 'Total')
    # worksheet.write(row, 1, '=SUM(B1:B4)')
    #
    # workbook.close()

    # fh = codecs.open("hello.txt", "r", "utf-8")
    # for line in fh:
    #     key, url, price, category = line.split("\t")
    #     print(line)
    # fh.close()
    #
    # global driver
    # driver = webdriver.Chrome()
    # # driver.get("file:///Users/lane/Desktop/test.html")
    # driver.get("file:///Users/lane/Desktop/second.html")
    # update_category_to_runqiang()
    #
    # driver.execute_script("window.open(arguments[0])", "https://www.baidu.com")
    # driver.switch_to.window(driver.window_handles[2])
    # driver.close()
    # driver.switch_to.window(driver.window_handles[1])
    # driver.close()
    pass


def choose_runqiang():
    driver.implicitly_wait(20)
    runqiang = driver.find_element_by_xpath("//*[@id='main']/div/div/div[1]/div[1]/div[2]/div/ul/li[1]/ul/li[1]/a")
    runqiang.send_keys(Keys.ENTER)


def read_input_excel():
    wb = xlrd.open_workbook('/Users/lane/Desktop/input.xlsx')
    sheet = wb.sheet_by_index(0)
    dict = {}
    for row in range(3, sheet.nrows):
        dict[sheet.cell(row, 1).value] = sheet.cell(row, 7).value
    return dict


def read_input_excel():
    wb = xlrd.open_workbook('/Users/lane/Desktop/input.xlsx')
    sheet = wb.sheet_by_index(0)
    list = []
    for row in range(3, sheet.nrows):
        list.append((sheet.cell(row, 1).value, sheet.cell(row, 7).value))
    return list


def sort_products_by_key():
    driver.implicitly_wait(20)
    runqiang = driver.find_element_by_xpath(
        "//*[@id='main']/div/div/div[2]/div/div[2]/div[2]/div/table/thead/tr/th[3]/a")
    runqiang.send_keys(Keys.ENTER)


def click_page(page_idx):
    driver.implicitly_wait(20)
    runqiang = driver.find_element_by_xpath(
        "//*[@id='main']/div/div/div[2]/div/div[2]/div[3]/ul/li[" + str(page_idx) + "]/a")
    runqiang.send_keys(Keys.ENTER)


def get_page():
    list = []
    try:
        for page_idx in range(1, 16):
            driver.implicitly_wait(20)
            path = "//*[@id='main']/div/div/div[2]/div/div[2]/div[2]/div/table/tbody/tr[" + str(
                page_idx) + "]/td[2]/a[2]"
            url = driver.find_element_by_xpath(path).get_attribute("href")

            path = "//*[@id='main']/div/div/div[2]/div/div[2]/div[2]/div/table/tbody/tr[" + str(page_idx) + "]/td[3]"
            driver.implicitly_wait(20)
            key = driver.find_element_by_xpath(path).text
            list.append((key, url))

    finally:
        return list


def set_search_option_by_key():
    driver.implicitly_wait(20)
    select = driver.find_element_by_xpath("//*[@id='main']/div/div/div[2]/div/div[2]/div[1]/div/div/span/span[1]/span")
    select.send_keys(Keys.ENTER)
    driver.implicitly_wait(20)
    input = driver.find_element_by_xpath("//*[@id='products_index']/span/span/span[1]/input")
    input.send_keys(u"产品编号")
    input.send_keys(Keys.ENTER)


def search_product_by_key(key):
    driver.implicitly_wait(20)
    input = driver.find_element_by_xpath("//*[@id='main']/div/div/div[2]/div/div[2]/div[1]/div/div/input")
    time.sleep(1)
    input.clear()
    input.send_keys(key)
    input.send_keys(Keys.ENTER)


def get_products_onebyone():
    driver.implicitly_wait(100)
    url = driver.find_element_by_xpath(
        "//*[@id='main']/div/div/div[2]/div/div[2]/div[2]/div/table/tbody/tr[1]/td[2]/a[2]").get_attribute(
        "href")
    category = driver.find_element_by_xpath(
        "//*[@id='main']/div/div/div[2]/div/div[2]/div[2]/div/table/tbody/tr/td[6]").text
    return url, category


def get_products_pagebypage():
    key_price_dict = read_input_excel()
    choose_runqiang()
    sort_products_by_key()

    list = [6, 9, 10, 11, 12]
    for page_idx in range(9, 100):
        list.append(12)
    list.extend([11, 10, 9, 1])

    results = []
    for page_idx in range(0, len(list)):
        key_url = get_page()
        print(page_idx + 1)
        for key, url in key_url:
            if key in key_price_dict:
                results.append((key, url, key_price_dict.get(key)))
        click_page(list[page_idx])
    file = open("/Users/lane/Desktop/allbyall", "w")
    for key, url, price in results:
        file.writelines(str(key) + " " + str(url) + " " + str(price) + "\n")
    file.close()
    return results


def update_category_to_whatever(name):
    search_name = "润强机电"
    if "活塞" in name:
        search_name = "润强活塞"
    elif "导向环" in name:
        search_name = "润强活塞"
    elif "油缸" in name:
        search_name = "润强油缸"
    elif "密封包" in name:
        search_name = "润强密封包"
    elif "管" in name:
        search_name = "润强泵车"
    elif "修理包" in name:
        search_name = "润强修理包"
    elif "分动箱" in name:
        search_name = "润强分动箱"
    elif "滤芯" in name:
        search_name = "润强滤芯"
    elif "总成" in name:
        search_name = "润强总成"
    elif "油泵" in name:
        search_name = "润强油泵"
    elif "切割环" in name:
        search_name = "润强切割环"

    locator = (By.XPATH, "//*[@id='main']/div[1]/div/div[2]/form/div[1]/div[4]/div/span")
    WebDriverWait(driver, 200).until(EC.element_to_be_clickable(locator))
    time.sleep(3)
    driver.implicitly_wait(20)
    combobox = driver.find_element_by_xpath("//*[@id='main']/div[1]/div/div[2]/form/div[1]/div[4]/div/span")
    combobox.click()
    search_input = driver.find_element_by_xpath("//*[@id='products_edit']/span/span/span[1]/input")
    search_input.send_keys(search_name)
    search_input.send_keys(Keys.ENTER)


def update_category_to_runqiang():
    driver.implicitly_wait(20)
    combobox = driver.find_element_by_xpath("//*[@id='main']/div[1]/div/div[2]/form/div[1]/div[4]/div/span")
    locator = (By.XPATH, "//*[@id='main']/div[1]/div/div[2]/form/div[1]/div[4]/div/span")
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(locator))
    combobox.click()
    input = driver.find_element_by_xpath("//*[@id='products_edit']/span/span/span[1]/input")
    input.send_keys(u"润强机电")
    input.send_keys(Keys.ENTER)


def clear_purchase_price():
    driver.implicitly_wait(20)
    price_tab = driver.find_element_by_xpath("//*[@id='main']/div[1]/div/div[1]/ul/li[5]/a")
    price_tab.send_keys(Keys.ENTER)

    driver.implicitly_wait(20)
    purchase_price = driver.find_element_by_xpath("//*[@id='purchase-price-table']/tbody/tr/td[3]/input")
    purchase_price.clear()
    purchase_price.send_keys(str(0.0))


def update_sell_price(price):
    driver.implicitly_wait(20)
    price_tab = driver.find_element_by_xpath("//*[@id='main']/div[1]/div/div[1]/ul/li[5]/a")
    price_tab.send_keys(Keys.ENTER)

    driver.implicitly_wait(20)
    sell_price = driver.find_element_by_xpath("//*[@id='sale-price-table']/tbody/tr/td[3]/input")
    sell_price.clear()
    sell_price.send_keys(str(price))


def update_one_by_one(url, value):
    driver.execute_script("window.open(arguments[0])", url)
    driver.switch_to.window(driver.window_handles[1])

    update_category_to_whatever(value)
    clear_purchase_price()
    # update_category_to_runqiang()
    # update_sell_price(value)

    btn = driver.find_element_by_xpath("//*[@id='main']/div[1]/div/div[1]/div[2]/a[1]")
    btn.send_keys(Keys.RETURN)

    driver.close()
    driver.switch_to.window(driver.window_handles[0])


def get_infos():
    key_price = read_input_excel()
    set_search_option_by_key()
    results = []
    for key, price in key_price:
        search_product_by_key(key)
        url, category = get_products_onebyone()
        results.append((key, url, price, category))

    write_output_excel(results)


def read_excel(name):
    wb = xlrd.open_workbook(name)
    sheet = wb.sheet_by_index(0)
    results = []
    for row in range(1082, sheet.nrows):
        result = []
        for col in range(sheet.ncols):
            result.append(sheet.cell(row, col).value)
        results.append(result)
    return results


if __name__ == "__main__":
    # test()
    # get_infos()

    results = read_excel('products.xlsx')
    login()
    set_search_option_by_key()
    for result in results:
        print(result)
        key = result[0]
        search_product_by_key(key)
        url, category = get_products_onebyone()
        update_one_by_one(url, result[1])
    # get_products_pagebypage()

    driver.close()
