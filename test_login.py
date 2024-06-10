import sys
import os
sys.path.append("C:/Users/vkarthikesan/PycharmProjects/Python_Selenium_task/Task27")
import pytest
from selenium import webdriver
from login_page import LoginPage
from XLU import XLU
from datetime import datetime

@pytest.fixture(scope="module")
def setup():
    driver = webdriver.Chrome()
    driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")
    yield driver
    driver.quit()

def test_login(setup):
    driver = setup
    file_path = "LoginTestData.xlsx"
    sheet_name = "Sheet1"
    row_count = XLU.get_row_count(file_path, sheet_name)

    for row_num in range(2, row_count + 1):
        test_id = XLU.read_data(file_path, sheet_name, row_num, 1)
        username = XLU.read_data(file_path, sheet_name, row_num, 2)
        password = XLU.read_data(file_path, sheet_name, row_num, 3)
        date_time = datetime.now()

        print(f"Starting test for row {row_num} with username: {username}")
        login_page = LoginPage(driver)
        login_page.enter_username(username)
        login_page.enter_password(password)
        login_page.click_login()

        if login_page.is_login_successful():
            result = "Passed"
            XLU.fill_green_color(file_path, sheet_name, row_num, 7)
        else:
            result = "Failed"
            XLU.fill_red_color(file_path, sheet_name, row_num, 7)

        XLU.write_data(file_path, sheet_name, row_num, 4, date_time.date())
        XLU.write_data(file_path, sheet_name, row_num, 5, date_time.time())
        XLU.write_data(file_path, sheet_name, row_num, 6, "Vikram")
        XLU.write_data(file_path, sheet_name, row_num, 7, result)

        # Reset the login page
        driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")
