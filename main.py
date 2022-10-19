from selenium import webdriver
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os
import xlsxwriter
import openpyxl

""" 
This was written for getting product information in order to help business pricing process.
"""

browser = webdriver.Chrome()

browser.get("http://erkoloyuncak.com.tr/")

# to close ad pop-up
close_popup = browser.find_element_by_xpath("/html/body/form/div[3]/div/div[1]/section/div[3]/div[1]/div/div/div/button")

close_popup.click()


login_button = browser.find_element_by_xpath("/html/body/form/div[3]/div/header/div/div[1]/div/nav[3]/ul/li[1]/a")

login_button.click()


time.sleep(2)

# to get user info and login elements
user_id = browser.find_element_by_xpath("/html/body/form/div[3]/div/div[1]/div/div[2]/div/div[1]/div[1]/input")
user_pw = browser.find_element_by_xpath("/html/body/form/div[3]/div/div[1]/div/div[2]/div/div[1]/div[2]/input")
user_login = browser.find_element_by_xpath("/html/body/form/div[3]/div/div[1]/div/div[2]/div/div[1]/div[4]/span/input")

print('Enter your ID and pass')
id_input = input()
pw_input = input()

# send id and password and make it login
user_id.send_keys(id_input)
user_pw.send_keys(pw_input)

user_login.click()

time.sleep(3)

# website links are pre-defined
links = open("links.txt" , encoding= 'utf-8').read().splitlines()

# to gather data
for link in links:
    try:
        browser.get(link)

        products = WebDriverWait(browser, 5).until(EC.visibility_of_all_elements_located((By.XPATH, "//div[@class='col-lg-4 col-md-4 col-sm-4 product']//div[@class='product-info']//h5//a")))

        for product in products:
            # get href
            href = product.get_attribute('href')

            # open new window with specific href
            browser.execute_script("window.open('" +href +"');")
            # switch to new window
            browser.switch_to.window(browser.window_handles[1])

            time.sleep(1)
            # gather product elements
            product_name  = browser.find_element_by_id("ContentPlaceHolder1_body_lblProductName")
            product_barcode = browser.find_element_by_id("ContentPlaceHolder1_body_lblBarkod")
            product_code = browser.find_element_by_id("ContentPlaceHolder1_body_lblUrunKodu")
            product_price = browser.find_element_by_id("ContentPlaceHolder1_body_lblFiyatNakit")

            # write down the product information into txt file
            f = open("productname.txt", "a+")
            f.write(product_name.text + "\n")
            f.close()

            f = open("productbarcode.txt", "a+")
            f.write(product_barcode.text + "\n")
            f.close()

            f = open("productcode.txt", "a+")
            f.write(product_code.text + "\n")
            f.close()

            f = open("productprice.txt", "a+")
            f.write(product_price.text + "\n")
            f.close()

            browser.close()
            # back to main window
            browser.switch_to.window(browser.window_handles[0])
    except:
        pass

browser.close()


# data cleansing
f = open("productname.txt", "r", encoding= 'utf-8', errors='ignore')

for line in f:
    words = line.split()
    f = open("productcode.txt", "a+")
    f.write(words[0] + "\n")
    f.close()
    words = line.split(None, 1)[1]
    f = open("productname.txt", "a+")
    f.write(words)
    f.close()

f.close()

# create main excel file
workbook = xlsxwriter.Workbook('ProductList.xlsx')
worksheet = workbook.add_worksheet()

# start from the first cell
# rows and columns are zero indexed
row = 0
column = 0

content_name = open("productname.txt", "r+", encoding="utf-8", errors='ignore')

# iterating through content list
for item in content_name:
    # write operation perform
    worksheet.write(row, column, item)

    # incrementing the value of row by one
    # with each iteratons
    row += 1

row = 0
column += 1


# to write down all information from txt files to excel file
content_code = open("productcode.txt", "r+", encoding="utf-8", errors='ignore')

for item in content_code:

    worksheet.write(row, column, item)

    row += 1

row = 0
column += 1



content_barcode = open("productbarcode.txt", "r+", encoding="utf-8", errors='ignore')

for item in content_barcode:
    worksheet.write(row, column, item)

    row += 1


row = 0
column += 1



content_price = open("productprice.txt", "r+", encoding="utf-8", errors='ignore')

for item in content_price:
    worksheet.write(row, column, item)

    row += 1

# setting columns
worksheet.set_column('A:A', 45)
worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 10)

workbook.close()

# clean unnecessary files
try:
    os.remove("productprice.txt")
except:
    pass
try:
    os.remove("productname.txt")
except:
    pass
try:
    os.remove("productbarcode.txt")
except:
    pass
try:
    os.remove("productcode.txt")
except:
    pass