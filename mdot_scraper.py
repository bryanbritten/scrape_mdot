# This first block of code just ensures that the user has all of the necessary packages installed
import subprocess
import sys

print("[+] Ensuring you have all of the necessary packages installed. You can review the requirements.txt file to see which packages are being installed.")
subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt", "--quiet"])
print("[+] Required libraries have been installed.")


import time
from datetime import datetime
from io import StringIO
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


def import_project_numbers():
    fpath = Path(input("Please provide the filepath to your list of project numbers: "))
    while True:
        if not fpath.exists():
            print("[-] That filepath does not exist.")
            fpath = Path(input("Please try again (press CTRL + C to exit): "))
        elif fpath.is_dir():
            print("[-] The filepath must point to a file, not a folder.")
            fpath = Path(input("Please try again (press CTRL + C to exit): "))
        elif fpath.suffix not in [".csv", ".txt"]:
            print("[-] Project numbers must be stored in a CSV or TXT file.")
            fpath = Path(input("Please try again (press CTRL + C to exit): "))
        else:
            break

    with open(fpath, "r") as f:
        project_numbers = f.readlines()

    return map(str.strip, project_numbers)


def set_output_directory():
    fpath = Path(input("Please provide a folder where the data should be saved: "))
    while True:
        if not fpath.exists():
            print("[-] That folder does not exist.")
            fpath = Path(input("Please try again (press CTRL + C to exit): "))
        elif not fpath.is_dir():
            print("[-] That file path doesn't seem to point to a directory.")
            fpath = Path(input("Please try again (press CTRL + C to exit): "))
        else:
            break

    return str(fpath).rstrip("/")


def get_chrome_driver():
    service = ChromeService(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    driver = webdriver.Chrome(service=service)
    return driver


def get_number_of_pages(driver):
    # not all contract numbers have multiple pages
    try:
        navigation = driver.find_element(By.CLASS_NAME, "page-navigation")
        links = navigation.find_elements(By.TAG_NAME, "a")
        return len(links) - 4
    except NoSuchElementException:
        return 1


def get_next_button(driver):
    try:
        navigation = driver.find_element(By.CLASS_NAME, "page-navigation")
        links = navigation.find_elements(By.TAG_NAME, "a")
        return links[-2]
    except NoSuchElementException:
        return None


def get_home_page(driver):
    BASE_URL = 'https://mdotjboss.state.mi.us/CCI/'
    driver.get(BASE_URL)

    # make sure that the inputs load so we know the page loaded successfully
    WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, 'selectedReportType'))
    )
    
    return driver


def select_from_dropdown(driver, dropdown_id, value):
    dropdown = Select(driver.find_element(By.ID, 'selectedReportType'))
    dropdown.select_by_visible_text(value)


def enter_project_number_in_form(driver, project_number):
    contract_number_input = driver.find_element(By.ID, 'contractProjectNum')
    contract_number_input.clear()
    contract_number_input.send_keys(project_number)
    contract_number_input.send_keys(Keys.ENTER)


def parse_subcontract_data(driver):
    data = []

    # Not every project has subcontracts associated with it.
    # If this project does not, then bail out.
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "subContractTable"))
        )
    except TimeoutException:
        return data

    num_pages = get_number_of_pages(driver)
    for _ in range(num_pages):
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "subcontrtactsId"))
        )

        table = driver.find_element(By.ID, "subContractTable")
        contractor_name = driver.find_element(
                By.XPATH, 
                '//div[@id="subcontrtactsId"]/div[@class="panel-body"]/div[@class="row"][1]/div[2]'
        ).text
        table_text = table.get_attribute('outerHTML')
        table_strIO = StringIO(str(table_text))
        subcontract_data = pd.read_html(table_strIO)[0]
        subcontract_data['Prime Contractor'] = contractor_name
        data.append(subcontract_data)

        next_button = get_next_button(driver)
        if next_button is not None:
            next_button.click()
    
    return data


def get_orig_contract_amount(driver):
    driver = get_home_page(driver)
    select_from_dropdown(driver, 'selectedReportType', 'General Contract Level Information')
    enter_project_number_in_form(driver, project_number)

    orig_contract_amount = driver.find_element(By.XPATH, '//div[@id="federalDivId"]/div[10]/div[2]').text
    return orig_contract_amount


def extract_project_data(driver, project_number):
    driver = get_home_page(driver)

    select_from_dropdown(driver, 'selectedReportType', 'Subcontracts')
    enter_project_number_in_form(driver, project_number)

    data = parse_subcontract_data(driver)
    df = pd.concat(data, ignore_index=True, sort=False)

    orig_contract_amount = get_orig_contract_amount(driver)
    df['Orig. Contract Amt'] = orig_contract_amount

    df['Project Number'] = project_number
    return df


def clean_data(data):
    data['Orig. Contract Amt'] = data['Orig. Contract Amt'].str.replace('$', '').str.replace(',', '').astype(float)
    data['SubCont  Value'] = data['SubCont  Value'].str.replace('$', '').str.replace(',', '').astype(float)
    data.columns = [' '.join(column.split()) for column in data.columns]
    return data


if __name__ == "__main__":
    project_numbers = import_project_numbers()
    output_directory = set_output_directory()

    # only support Chrome for the moment
    driver = get_chrome_driver()

    failed_projects = []
    for project_number in project_numbers:
        data = extract_project_data(driver, project_number)
        if not data.empty:
            fpath = f"{output_directory}/{datetime.today().strftime('%Y%m%d')}.csv"
            data = clean_data(data)
            data.to_csv(fpath, index=False, mode="a")
        else:
            failed_projects.append(project_number)
        time.sleep(3)

    if failed_projects:
        print(
            "[-] The following projects did not appear to have subcontracts associated with them:"
        )
        for project in failed_projects:
            print(project)
