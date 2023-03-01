import time
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import keyboard
from openpyxl import load_workbook

import xpath_file as xpath

counter = 0


def function(driver, find_by, find_by_value, operation_to_do, success_message, failed_message, send_key_value='', index_value=0, find_time_parameter=20):
    global counter

    try:
        if find_by == 'xpath':
            element = WebDriverWait(driver, find_time_parameter).until(EC.presence_of_element_located((By.XPATH, find_by_value)))
        elif find_by == 'id':
            element = WebDriverWait(driver, find_time_parameter).until(EC.presence_of_element_located((By.ID, find_by_value)))
        elif find_by == 'partial_link_text':
            element = WebDriverWait(driver, find_time_parameter).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, find_by_value)))
        elif find_by == 'link_text':
            element = WebDriverWait(driver, find_time_parameter).until(EC.presence_of_element_located((By.LINK_TEXT, find_by_value)))
        elif find_by == 'select':
            element = Select(driver.find_element('xpath', find_by_value))

        if operation_to_do == 'click':
            element.click()
        elif operation_to_do == 'send_keys':
            element.send_keys(send_key_value)
        elif operation_to_do == 'select_by_index':
            element.select_by_index(index_value)

    except Exception:
        if counter < 2:
            print('I will try again')
            counter += 1
            time.sleep(2)
            function(driver, find_by, find_by_value, operation_to_do, success_message, failed_message, send_key_value, index_value, find_time_parameter)
        else:
            counter = 0
            print(failed_message)
            keyboard.wait("Enter")
    else:
        counter = 0
        print(success_message)


def function_with_wait(driver, find_by, find_by_value, second_element, operation_to_do, success_message, failed_message, send_key_value='', find_time_parameter=20):
    global counter

    try:
        if find_by == 'xpath':
            element = WebDriverWait(driver, find_time_parameter).until(
                EC.presence_of_element_located((By.XPATH, find_by_value)))
        elif find_by == 'id':
            element = WebDriverWait(driver, find_time_parameter).until(
                EC.presence_of_element_located((By.ID, find_by_value)))
        elif find_by == 'partial_link_text':
            element = WebDriverWait(driver, find_time_parameter).until(
                EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, find_by_value)))
        elif find_by == 'link_text':
            element = WebDriverWait(driver, find_time_parameter).until(
                EC.presence_of_element_located((By.LINK_TEXT, find_by_value)))

        if operation_to_do == 'click':
            element.click()
        elif operation_to_do == 'send_keys':
            element.send_keys(send_key_value)

        WebDriverWait(driver, find_time_parameter).until(
            EC.presence_of_element_located((By.XPATH, second_element)))

    except Exception:
        if counter < 2:
            print('I will try again')
            counter += 1
            time.sleep(2)
            function_with_wait(driver, find_by, find_by_value, second_element, operation_to_do, success_message, failed_message, send_key_value, find_time_parameter)
        else:
            counter = 0
            print(failed_message)
            keyboard.wait("Enter")
    else:
        counter = 0
        print(success_message)


def open_url(driver):
    try:
        driver.get(xpath.url)
    except Exception as e:
        print('Chyba pri otvarani url. Error: ', e)
    else:
        print('URL.......................................OK')


def write_to_excel(url, name, iteration):
    sting = str(iteration)
    try:
        outputs_path = r'outputs.xlsx'

        wb = load_workbook(outputs_path)

        work_sheet = wb.active
        work_sheet['A'+sting] = name
        work_sheet['B'+sting] = url

        wb.save("outputs.xlsx")

    except Exception as e:
        print('Chyba pri write to excel Error: ', e)
    else:
        print('Write to excel.................OK')
        print('')
