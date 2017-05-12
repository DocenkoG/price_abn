from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import unittest, time, re
import os
https://b2b.abn.ru/get_price_list.php?action=proccess&login=72392&reqstr=-1_6701,6575,6582,6630,6710,6584,7044,6608,6657,7047,6638,6618,6619,6681,6639,7365,6621,6641,6622,7081,6589,6660,6610,6624,6591,7369,6625,6684,6578,6696,6627,7413,6665,6667,6646,7410,6596,7199,6671,6658,6672,6598,6704,6698,7366,6649,6632,6687,6651,7364,6690,6614,6615,6602,6616,6635,7352,6708,6712,7378,6599,7348,6692,7362_1_1_1_0_1&fileMode=xls
class Abn(unittest.TestCase):
    def setUp(self):
        ffprofile = webdriver.FirefoxProfile()
        ffprofile.set_preference("browser.download.dir", os.getcwd()+'\\tmp')
        ffprofile.set_preference("browser.download.folderList",2);
        ffprofile.set_preference("browser.helperApps.neverAsk.saveToDisk","application/xls,application/octet-stream,application/vnd.ms-excel,application/x-excel,application/x-msexcel,application/excel")
        self.driver = webdriver.Firefox(ffprofile)
        self.driver.implicitly_wait(30)
        self.base_url = "https://b2b.abn.ru/"
        self.verificationErrors = []
        self.accept_next_alert = True
    
    def test_abn(self):
        driver = self.driver
        driver.get(self.base_url + "/get_price_list.php?action=proccess&login=72392&reqstr=-1_6701,6575,6582,6630,6710,6584,7044,6608,6657,7047,6638,6618,6619,6681,6639,7365,6621,6641,6622,7081,6589,6660,6610,6624,6591,7369,6625,6684,6578,6696,6627,7413,6665,6667,6646,7410,6596,7199,6671,6658,6672,6598,6704,6698,7366,6649,6632,6687,6651,7364,6690,6614,6615,6602,6616,6635,7352,6708,6712,7378,6599,7348,6692,7362_1_1_1_0_1&fileMode=xls")
        driver.find_element_by_css_selector("div.image").click()
        driver.find_element_by_css_selector("a.download_ico.xls").click()
        time.sleep(20)

    def is_element_present(self, how, what):
        try: self.driver.find_element(by=how, value=what)
        except NoSuchElementException: return False
        return True
    
    def is_alert_present(self):
        try: self.driver.switch_to_alert()
        except NoAlertPresentException: return False
        return True
    
    def close_alert_and_get_its_text(self):
        try:
            alert = self.driver.switch_to_alert()
            alert_text = alert.text
            if self.accept_next_alert:
                alert.accept()
            else:
                alert.dismiss()
            return alert_text
        finally: self.accept_next_alert = True
    
    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)

if __name__ == "__main__":
    unittest.main()
