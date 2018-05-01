from selenium import webdriver
from bs4 import BeautifulSoup
import time
import xlrd
import xlwt
import re

class Spider:

    def __init__(self):
        self.driver = webdriver.Chrome()
        self.keyWords = []
        self.resultDic = {}
        # TODO 作用域可能有问题，值可能会不对
        date = time.strftime("%Y-%m-%d %H_%M_%S", time.localtime())
        self.resultExcelName = date + 'position.xls'

    def readKeyWord(self, filePath):
        global keyWords
        excel = xlrd.open_workbook(filePath.get())
        table = excel.sheet_by_name(u'key_word')
        keyWords = table.col_values(0)


    def searchKeyWord(self):
        for keyWord in keyWords:
            data = {
                keyWord: {keyWord + '_ad_position': [], keyWord + '_natural_position': [], keyWord + '_all_position': []}}
            self.resultDic.update(data)
            self.driver.get("https://www.amazon.co.uk/")
            time.sleep(10)
            self.driver.find_element_by_id("twotabsearchtextbox").click()
            self.driver.find_element_by_id("twotabsearchtextbox").clear()
            self.driver.find_element_by_id("twotabsearchtextbox").send_keys(keyWord)
            # 第一页
            self.driver.find_element_by_name("site-search").submit()
            time.sleep(10)
            self.readPage(keyWord)
            # 第2-10页
            for i in range(2, 11):
                self.driver.find_element_by_id("pagnNextString").click()
                time.sleep(10)

                # 读取页面信息
                self.readPage(keyWord)
        self.driver.close()


    def readPage(self, keyWord):
        html = self.driver.page_source
        soup = BeautifulSoup(html, features='lxml')
        all_result = soup.find_all('h2')

        for result in all_result:
            self.resultDic.get(keyWord).get(keyWord + '_all_position').append(result.get_text())

            if re.search('Sponsored', result.get_text()) is None:
                self.resultDic.get(keyWord).get(keyWord + '_natural_position').append(result.get_text())
                continue

            self.resultDic.get(keyWord).get(keyWord + '_ad_position').append(result.get_text())


    def createResultExcel(self, outputFilePath):
        workbook = xlwt.Workbook(encoding='ascii')

        keyWordNum = 1

        for keyWord in keyWords:
            worksheet = workbook.add_sheet(str(keyWordNum) + '_ad_position')
            row = 0
            for title in self.resultDic.get(keyWord).get(keyWord + '_ad_position'):
                worksheet.write(row, 0, title)
                row += 1

            worksheet = workbook.add_sheet(str(keyWordNum) + '_natural_position')
            row = 0
            for title in self.resultDic.get(keyWord).get(keyWord + '_natural_position'):
                worksheet.write(row, 0, title)
                row += 1

            worksheet = workbook.add_sheet(str(keyWordNum) + '_all_position')
            row = 0
            for title in self.resultDic.get(keyWord).get(keyWord + '_all_position'):
                worksheet.write(row, 0, title)
                row += 1

            keyWordNum += 1

        workbook.save(outputFilePath.get()+'/'+self.resultExcelName)
