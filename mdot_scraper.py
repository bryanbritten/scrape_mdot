import time
import sys
from pathlib import Path
from winreg import HKEY_CURRENT_USER, OpenKey, QueryValueEx


def install_library(lib_name) -> None:
    import subprocess

    subprocess.check_call([sys.executable, "-m", "pip", "install", lib_name])
    print(
        f"[+] {lib_name}, a required library, wasn't installed so it has been installed for you."
    )


try:
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import Select
    from selenium.webdriver.support.wait import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import NoSuchElementException
except ImportError:
    install_library("selenium")
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import Select
    from selenium.webdriver.support.wait import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import NoSuchElementException


try:
    import pandas as pd
except ModuleNotFoundError:
    install_library("pandas")
    install_library("lxml")
    import pandas as pd


# stolen from https://stackoverflow.com/questions/19037216/how-to-get-a-name-of-default-browser-using-python
def get_default_browser() -> str:
    with OpenKey(
        HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice",
    ) as regkey:
        browser = QueryValueEx(regkey, "ProgId")[0]

    return browser


def import_project_numbers() -> list[str]:
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


def get_contract_data(driver, project_number):
    BASE_URL = "https://mdotjboss.state.mi.us/CCI/"
    driver.get(BASE_URL)

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "selectedReportType"))
    )
    report_type = Select(driver.find_element(By.ID, "selectedReportType"))
    report_type.select_by_visible_text("Subcontracts")

    contract_number = driver.find_element(By.ID, "contractProjectNum")
    contract_number.clear()
    contract_number.send_keys(project_number)
    contract_number.send_keys(Keys.ENTER)

    data = []

    # Wait for the table to show up before trying to get the number of pages
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "subContractTable"))
    )

    num_pages = get_number_of_pages(driver)
    for _ in range(num_pages):
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "subcontrtactsId"))
        )

        table = driver.find_element(By.ID, "subContractTable")
        data.append(pd.read_html(table.get_attribute("outerHTML"))[0])

        next_button = get_next_button(driver)
        if next_button is not None:
            next_button.click()

    df = pd.concat(data, ignore_index=True, sort=False)
    return df


if __name__ == "__main__":
    project_numbers = import_project_numbers()
    output_directory = set_output_directory()

    # default_browser = get_default_browser()
    # only support Chrome for the moment
    default_browser = "ChromeHTML"
    if default_browser == "ChromeHTML":
        from selenium.webdriver.chrome.service import Service as ChromeService

        try:
            from webdriver_manager.chrome import ChromeDriverManager
        except ImportError:
            install_library("webdriver-manager")
            from webdriver_manager.chrome import ChromeDriverManager

        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
    else:
        print("This script can only run in Chrome. Please change your default browser.")
        sys.exit(0)

    for project_number in project_numbers:
        data = get_contract_data(driver, project_number)
        fpath = f"{output_directory}/{project_number}.csv"
        data.to_csv(fpath, index=False)
        time.sleep(3)
