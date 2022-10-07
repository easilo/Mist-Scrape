# --------------------------------------------------------------------------
# Author:           Erin Asilo
# Create date:      9/23/2022
# Ogranization:     RingCentral
# Email:            erin.asilo@ringcentral.com
#
# Purpose:          To automate the process of grabbing site client
#                   metrics from Mist and put them into a Google Sheet
# --------------------------------------------------------------------------

# --------------------------------------------------------------------------
# Imports
# --------------------------------------------------------------------------

from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import StaleElementReferenceException
import chromedriver_autoinstaller
import time
import os
from dotenv import load_dotenv
from Google import Create_Service
from datetime import date
from datetime import datetime
import sys
import gspread

# --------------------------------------------------------------------------
# Load .env file
# --------------------------------------------------------------------------

load_dotenv()
EMAIL = os.environ.get("EMAIL")
PASSWORD = os.environ.get("PASSWORD")

DOWNLOAD_PATH = os.environ.get("DOWNLOAD_PATH")
ZIP_PATH = os.environ.get("ZIP_PATH")
SECRET_PATH = os.environ.get("SECRET_PATH")
SERVICE_ACCOUNT = os.environ.get("SERVICE_ACCOUNT")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")

# --------------------------------------------------------------------------
# Create service with Google API
# --------------------------------------------------------------------------

gc = gspread.service_account(SERVICE_ACCOUNT)
spreadsheet = gc.open_by_key(SPREADSHEET_ID)

CLIENT_SECRET_FILE = SECRET_PATH
API_NAME = "sheets"
API_VERSION = "v4"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

spreadsheet_id = SPREADSHEET_ID
mySpreadsheets = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

# --------------------------------------------------------------------------
# Selenium chromedriver options
# --------------------------------------------------------------------------

chromedriver_autoinstaller.install()

options = webdriver.ChromeOptions()
# options.headless = True
options.add_argument("--window-size=1920,1080")
options.add_argument("start-maximized")
options.add_argument("--disable-gpu")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36"
)
# options.add_argument("--remote-debugging-port=9222")
options.add_argument("--no-sandbox")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--lang=en_US")

prefs = {"download.default_directory": DOWNLOAD_PATH}

options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)

wait = WebDriverWait(driver, 45)
action = ActionChains(driver)

sys_cls_clear = "Cls"

# --------------------------------------------------------------------------
# Date and time format
# --------------------------------------------------------------------------

today = datetime.now().strftime("%m/%d/%Y")
day = date.today().strftime("%A")

# --------------------------------------------------------------------------
# Function:     scrape()
# Purpose:      Log in to Mist and download the .csv file containing metrics
# --------------------------------------------------------------------------


def scrape():
    try:
        # Log into Mist
        print("Executing Mist Scrape ... ")
        print("    Logging into Mist ... ")
        driver.get("https://manage.mist.com")
        time.sleep(1.5)
        email = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "/html/body/div/div/div[2]/div[4]/div/div[1]/div[2]/form/div[1]/div/label/input",
                )
            )
        )
        email.send_keys(EMAIL)
        login = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Next')]"))
        )
        login.click()
        time.sleep(3)
        password = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//*[@id="app-content"]/div/div[2]/div[4]/div/div[1]/div[2]/form/div[2]/div/label/input',
                )
            )
        )
        password.send_keys(PASSWORD)
        login = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Sign In')]"))
        )
        login.click()
        time.sleep(2)
        wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//*[@id="app-body"]/div/div[2]/div[2]/div[1]/div[2]/div/div/span/span/button[1]',
                )
            )
        )
        print("    Grabbing metrics ...\n")
        time.sleep(4)
        analytics = driver.find_element(
            By.XPATH,
            '//*[@id="app-body"]/div/div[2]/div[2]/div[1]/div[2]/div/div/span/span/button[1]',
        )
        analytics.click()
        time.sleep(3)
        dropdown = wait.until(
            EC.element_to_be_clickable(
                (
                    By.CLASS_NAME,
                    "timeRangeSelector",
                )
            )
        )
        dropdown.click()
        timeframe = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//*[@id="app-body"]/div/div[2]/div[2]/div[1]/div[2]/div/div[2]/div/div/div/div/div[1]/div[3]',
                )
            )
        )
        timeframe.click()
        time.sleep(5)
        dropdown = wait.until(
            EC.element_to_be_clickable(
                (
                    By.CLASS_NAME,
                    "scopeSelector",
                )
            )
        )
        dropdown.click()
        site_finder = driver.find_element(By.CLASS_NAME, "entityScroll")
        sites_elements = site_finder.find_elements(By.TAG_NAME, "li")
        sites = []
        values = []
        for items in sites_elements:
            if len(items.text) <= 3:
                sites.append(items.text)
        """
        site_finder = driver.find_element(
            By.CLASS_NAME,
            "/html/body/div[1]/div/div[2]/div[2]/div[1]/div[2]/div/div[1]/div/div/div[2]/div/div[1]/input",
        )
        
        connect_element = driver.find_element(
            By.XPATH,
            "//*[contains(text(), 'Successful Connects')]/../following-sibling::div[1]/p",
        )"""
        for site in sites:
            dropdown.click()
            time.sleep(2)
            staleElement = True
            while staleElement:
                try:
                    site_finder = driver.find_element(
                        By.XPATH,
                        "/html/body/div[1]/div/div[2]/div[2]/div[1]/div[2]/div/div[1]/div/div/div[2]/div/div[1]/input",
                    )
                    staleElement = True
                except StaleElementReferenceException as e:
                    staleElement = False
            time.sleep(2)
            site_finder.send_keys(site)
            time.sleep(1)
            site_finder.submit()
            time.sleep(3)
            connect_element = driver.find_element(
                By.XPATH,
                "//*[contains(text(), 'Successful Connects')]/../following-sibling::div[1]/p",
            )
            values.append(connect_element.text)
            print("Site: " + site + "Successful Connects: " + connect_element.text())

        # print("Site: " + "Successful Connects: " + connect_element.text)

    except Exception as e:
        print(e)
        driver.quit()
        sys.stdout.flush()
        os.execv(
            sys.executable, [sys.executable, '"' + sys.argv[0] + '"'] + sys.argv[1:]
        )


# --------------------------------------------------------------------------
# Function:     parse_and_upload()
# Purpose:      Parses through .csv list and saves the data to a variable
#               per site, then sends data to Google sheet through API
# --------------------------------------------------------------------------


def parse_and_upload(data):
    # Test data ----- Comment out
    # data.extend([["TEST", "777"]])

    # Get titles of columns
    columns = spreadsheet.get_worksheet(1).row_values(1)
    columns_array = []
    title_array = []
    new_column = False
    for el in columns:
        title_array.append(el.split(", "))
        columns_array.append(el.split(", "))
    del columns_array[:2]
    # Prepare first 2 columns as the date and day
    values = [[today], [day]]
    for i in range(len(data)):
        sites = data[i][0]
        clients = data[i][1]
        site_array = [sites]
        # First line of data file is "sites" so skip this
        if sites == "Sites":
            continue
        # Check if the site from the data file is already in spreadsheet, if it is then match the data with the site
        if site_array in columns_array:
            columns_array[columns_array.index(site_array)][0] = clients
        # If site does not match any column in spreadsheet then add a new column
        else:
            title_array.append([sites])
            columns_array.extend([[clients]])
            new_column = True
        if sites != "Sites":
            print("Site: " + sites + " | Clients: " + clients)

    # Check if value is a string then erase it (for example if a site has 0 clients it will not show up in the list)
    for list in columns_array:
        for idx in range(len(list)):
            if list[idx] == today or list[idx] == day:
                continue
            elif list[idx][0].isalpha():
                list[idx] = " "
    values.extend(columns_array)

    print("\nUpdating Google sheet ...\n")
    worksheet_name = "Data!"
    cell_range_insert = "A2"

    value_range_body = {"majorDimension": "COLUMNS", "values": values}

    value_range_body1 = {"majorDimension": "COLUMNS", "values": title_array}

    # Update the column titles
    if new_column == True:
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            valueInputOption="USER_ENTERED",
            range=worksheet_name + "A1",
            body=value_range_body1,
        ).execute()

    # Update the values on spreadsheet
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        valueInputOption="USER_ENTERED",
        range=worksheet_name + cell_range_insert,
        body=value_range_body,
    ).execute()


# --------------------------------------------------------------------------
# Function:     main()
# Purpose:      Run the previous functions and catches errors
# --------------------------------------------------------------------------


def main():
    os.system(sys_cls_clear)
    try:
        print("Today is " + day + " " + today + "\n")
        scrape()
        # parse_and_upload(scrape())
        print("Done!\n")
    # Retry program if it fails
    except Exception as e:
        print(e)
        driver.quit()
        sys.stdout.flush()
        os.execv(
            sys.executable, [sys.executable, '"' + sys.argv[0] + '"'] + sys.argv[1:]
        )
    try:
        sys.exit(1)
    except:
        os._exit(1)


main()
