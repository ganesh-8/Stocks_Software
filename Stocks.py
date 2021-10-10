from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import os
import glob
import pandas as pd
import time
import numpy as np
import xlsxwriter

companyList = ["ABB", "Reliance Industries"]
browser = webdriver.Chrome(r'C:\\Users\\smart\\Downloads\\chromedriver_win32 (1)\\chromedriver.exe')

book = xlsxwriter.Workbook("C:/Users/smart/Desktop/AllStocksList.xlsx")
sheet = book.add_worksheet()
sheet.write(0,0,"CompanyName")
startYear = ['2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020']
endYear = ['2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021']
for i in range(10):
    sheet.write(0,i+1,"FY"+startYear[i]+"-"+endYear[i])
const_var = 1
for companyName in companyList:
    browser.get("https://www.bseindia.com/")
    searchBar = browser.find_element_by_id("getquotesearch")
    searchBar.send_keys(companyName)
    browser.implicitly_wait(10)
    searchBar.send_keys(Keys.ENTER)
    browser.implicitly_wait(50)
    archives = browser.find_element_by_xpath('//*[@id="getquoteheader"]/div[6]/div/div[3]/div/div[3]/div[2]/a/input')
    archives.send_keys(Keys.ENTER)

    browser.find_element_by_xpath('//*[@id="ContentPlaceHolder1_txtFromDate"]').click()
    browser.find_element_by_xpath('//*[@id="ui-datepicker-div"]/div/div/select[1]').send_keys("APR")
    browser.find_element_by_xpath('//*[@id="ui-datepicker-div"]/div/div/select[2]').send_keys("2011")
    browser.find_element_by_xpath('//*[@id="ui-datepicker-div"]/table/tbody/tr[1]/td[6]/a').send_keys("1")
    browser.find_element_by_xpath('//*[@id="ui-datepicker-div"]/table/tbody/tr[1]/td[6]/a').click()

    browser.implicitly_wait(50)

    browser.find_element_by_xpath(
        '/html/body/form/div[4]/div/div/div[1]/div/div[3]/div/div/table/tbody/tr[4]/td/table/tbody/tr/td/div/table/tbody/tr[1]/td[3]/input').click()
    browser.find_element_by_xpath('/html/body/div/div/div/select[1]').send_keys("MAR")
    browser.find_element_by_xpath('/html/body/div/div/div/select[1]').click()
    browser.implicitly_wait(10)
    browser.find_element_by_xpath('/html/body/div/table/tbody/tr[5]/td[4]/a').send_keys("31")
    browser.find_element_by_xpath('/html/body/div/table/tbody/tr[5]/td[4]/a').click()

    browser.implicitly_wait(10)
    browser.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnSubmit"]').click()

    browser.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnDownload1"]/i').click()

    time.sleep(10)

    list_of_files = glob.glob('C:/Users/smart/Downloads/*.csv')  # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    print(latest_file)

    data = pd.read_csv(latest_file)
    avgList = []
    df = pd.DataFrame(data, columns=["Date", "Close Price"])
    dates = data.loc[:, "Date"]
    df['Date'] = pd.to_datetime(df['Date'])

    startYear = ['2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020']
    endYear = ['2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021']
    closePriceAvgList = []

    i = 0

    while i < len(startYear):
        mask = ((df['Date'] >= startYear[i] + "-04-01") & (df['Date'] <= endYear[i] + "-03-31"))
        df1 = df.loc[mask]
        closePriceAvg = df1['Close Price'].mean()
        closePriceAvgList.append(closePriceAvg)
        i += 1
    i = 1
    for stock in closePriceAvgList:
        sheet.write(const_var, 0, companyName)
        sheet.write(const_var, i, stock)
        i += 1
        print(stock)
    const_var += 1

    print(df.loc[mask])
    print(closePriceAvg)
book.close()

