"""
Write the test scripts for performing the data driven testing
"""
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Import the data and locators
from TestLocators.locators import SauceLocators
from TestData.data import SauceData
from Utilities.excel_functions import ExcelFunctions

class TestSauceDDTF:

    def test_login_excel(self):
        self.excel_file = SauceData().excel_file
        self.sheet_number = SauceData().sheet_number

        self.excel = ExcelFunctions(self.excel_file, self.sheet_number)
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        
        self.driver.get(SauceData().url)
        self.driver.maximize_window()

        self.driver.implicitly_wait(10)

        # get the maximum row count inorder to loop properly
        self.rows = self.excel.row_count()

        # use for loop for fetching the datasets
        for row in range(2, self.rows+1):
            username = self.excel.read_data(row, 5)
            password = self.excel.read_data(row, 6)

            # to perform login with username and password from the excel file
            self.driver.find_element(by=By.ID, value=SauceLocators.username_locator).send_keys(username)
            self.driver.find_element(by=By.ID, value=SauceLocators.password_locator).send_keys(password)
            self.driver.find_element(by=By.ID, value=SauceLocators.loginbutton_locator).click()

            if SauceData().dashboard_url in self.driver.current_url:
                print("SUCCESS : Login Successful")
                self.excel.write_data(row, 7, "Test Pass")
                self.driver.back()

            elif SauceData().url in self.driver.current_url:
                print("FAIL : Login Failed")
                self.excel.write_data(row, 7, "Test Fail")
                self.driver.refresh()
            
        self.driver.quit()
