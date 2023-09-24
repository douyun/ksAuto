# coding=utf-8
import time

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By

ORDER_ID = 'KS-WH-JPE-236208'
# ORDER_ID = 'KS-WH-ZES-235481'


class KsAuto(object):

    def __init__(self):
        self.driver = webdriver.Edge(executable_path="D:\Tools\msedgedriver.exe")
        self.driver.maximize_window()
        self.driver.implicitly_wait(10)
        self.driver.get('https://www.knnns.net/webroot/decision/login?origin=af4835c8-c51f-465b-927f-a01edd5a01f3')
        self.action = ActionChains(self.driver)
        self.login_ks()

    def login_ks(self):
        self.driver.find_element(By.CSS_SELECTOR, 'input[placeholder=用户名]').send_keys('18192183575')
        self.driver.find_element(By.CSS_SELECTOR, 'input[placeholder=密码]').send_keys('mq880820')
        button = self.driver.find_element(By.CLASS_NAME, 'login-button')
        button.click()

    def to_common_page(self):
        for _ in range(3):
            try:
                folder1 = self.driver.find_element(By.XPATH, "//div[text()='分选/返工服务']")
                folder1.click()
                folder2 = self.driver.find_element(By.XPATH, ".//div[text()='通用']")
                folder2.click()
                self.driver.switch_to.frame('b5a1c566-7935-471a-9fa9-d3fc003f40db')
                return
            except:
                time.sleep(5)
        raise "to common page failed"

    def query_order(self, orderId):
        oderInputText = self.driver.find_element(By.XPATH, "//div[@widgetname='ORDERNUMBER']/div/input[@type='text']")
        oderInputText.clear()
        oderInputText.send_keys(orderId + '\n')
        time.sleep(1)
        oderInputText.send_keys(orderId + '\n')
        time.sleep(0.5)
        oderInputText.send_keys('\n')
        self.driver.find_element(By.XPATH, "//div[@widgetname='FORMSUBMIT']//button[text()='查询']").click()

    def _get_submit_button(self):
        return self.driver.find_element(By.XPATH, "//div[@widgetname='提交']")

    def _get_last_row_index(self):
        for _ in range(3):
            try:
                lastRow = self.driver.find_element(By.XPATH,
                                                 "//div[@id='content-container']//table//tr[last()-2]//button[text()='增']/../../../../../..")
                return int(lastRow.get_attribute('tridx'))
            except:
                time.sleep(1)
        raise


    def _double_submit(self):
        submitBtn = self._get_submit_button()
        submitBtn.click()
        time.sleep(3)
        submitBtn.click()
        time.sleep(4)

    def _single_submit(self):
        submitBtn = self._get_submit_button()
        submitBtn.click()
        time.sleep(4)

    def _is_exist_first_blank_line(self, lastLineNum):
        return lastLineNum == 9 and\
               len(self.driver.find_element(By.XPATH, "//div[@id='content-container']//table//td[@col='5'][@row='9']").text) == 0

    def _del_cell_old_data(self):
        row = self._get_last_row_index()
        for col in [2, 3, 4, 5, 9, 10, 12, 13]:
            cell = self.driver.find_element(By.XPATH, "//div[@id='content-container']//table//td[@col='{0}'][@row='{1}']".format(col, row))
            if len(cell.text) > 0:
                cell.click()
                self.driver.find_element(By.XPATH, "//div[contains(@class,'x-editor')]//input[@type='text']").clear()

    def wait_query_order_finished(self):
        for _ in range(2):
            try:
                self.driver.find_element(By.XPATH, "//div[@widgetname='FORMSUBMIT'][contains(@class,'ui-state-enabled')]//button[text()='查询']")
                break
            except:
                time.sleep(1)

    def _add_row(self):
        for _ in range(3):
            try:
                self.driver.find_element(By.XPATH, "//div[@id='content-container']//table//tr[last()-2]//button[text()='增']").click()
                break
            except:
                time.sleep(4)

    def add_new_row(self, rowCount=1):
        lastNumBeforeAdd = self._get_last_row_index()
        startShift = 1
        if self._is_exist_first_blank_line(lastNumBeforeAdd):
            rowCount -= 1
            startShift = 0
        for i in range(rowCount):
            self._add_row()
            sleepSeconds = 1 if lastNumBeforeAdd <= 100 else 5
            time.sleep(sleepSeconds)
            if i == 0:
                self._del_cell_old_data()
        if rowCount > 0:
            self._single_submit()
        lastNumAfterAdd = self._get_last_row_index()
        if lastNumAfterAdd - lastNumBeforeAdd == rowCount:
            return lastNumBeforeAdd + startShift, lastNumAfterAdd
        return -1, -1

    def update_row(self, row, col, value, isDelEnter=True):
        for _ in range(3):
            try:
                cell = self.driver.find_element(By.XPATH, "//div[@id='content-container']//table//td[@col='{1}'][@row='{0}']".format(row, col))
                cell.click()
                inputElement = self.driver.find_element(By.XPATH, "//div[contains(@class,'x-editor')]//input[@type='text']")
                break
            except:
                time.sleep(0.5)
        value = value.replace('\n', '') if isDelEnter else value
        inputElement.send_keys(value)

    def goto_expense_report(self):
        btn = self.driver.find_element(By.XPATH, "//div[@id='fr-btn-费用报表']//button[text()='费用报表']")
        btn.click()
        time.sleep(3)
        self.driver.find_element(By.XPATH, "//span[text()='确定']").click()
        time.sleep(3)

    def goto_quality_report(self):
        btn = self.driver.find_element(By.XPATH, "//div[@id='fr-btn-质量报表']//button[text()='质量报表']")
        btn.click()
        time.sleep(3)

    def submit_on_expense_page(self):
        btn = self.driver.find_element(By.XPATH, "//div[@widgetname='提交']//span[text()='提交']")
        for _ in range(2):
            btn.click()
            time.sleep(4)
            try:
                self.driver.find_element(By.XPATH, "//span[text()='确定']").click()
            except:
                pass
            time.sleep(1)

    def click_makesure_button(self):
        try:
            self.driver.find_element(By.XPATH, "//span[text()='确定']").click()
        except:
            pass

    def get_expense_row_element(self, countFromEnd):
        trElems = self.driver.find_elements(By.XPATH, "//div[@id='content-container']//table//tr")
        return trElems[-4-countFromEnd: -4]

    def refresh(self):
        for _ in range(3):
            try:
                self.driver.refresh()
                return
            except:
                time.sleep(5)
        raise 'refresh page failed'

    def refresh_by_click_button(self):
        self.driver.switch_to.parent_frame()
        commonBtn = self.driver.find_element(By.XPATH, "//div[contains(@class,'bi-basic-button')][text()='通用']")
        self.action.move_to_element(commonBtn).perform()
        refreshBtn = self.driver.find_element(By.XPATH, "//div[contains(@class,'bi-basic-button')]//div[text()='刷新']")
        refreshBtn.click()
        time.sleep(3)
        self.action.move_by_offset(0, 100).perform()
        self.driver.switch_to.frame('b5a1c566-7935-471a-9fa9-d3fc003f40db')



if __name__ == '__main__':
    ks = KsAuto()
    ks.to_common_page()
    ks.query_order(ORDER_ID)
    time.sleep(3)
    ks.refresh_by_click_button()
    # ks.update_row(23, 2, u'零件号')
    # ks.goto_expense_report()
    # ks.get_expense_row_element(2)
    # ks.submit_on_expense_page()


    # ks.goto_quality_report()

    # startIndex, endIndex = ks.add_new_row(3)
    # for i in range(startIndex, endIndex+1):
    #     ks.update_row(i, [])
    # ks._double_submit()
