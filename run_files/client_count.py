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
import chromedriver_autoinstaller
import time
import os
from dotenv import load_dotenv
import csv
from io import TextIOWrapper
from zipfile import ZipFile
from datetime import date
from datetime import datetime
import sys
import gspread
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2 import service_account
import re

# --------------------------------------------------------------------------
# Load .env file
# --------------------------------------------------------------------------

load_dotenv()
EMAIL = os.environ.get("EMAIL")
PASSWORD = os.environ.get("PASSWORD")

DOWNLOAD_PATH = os.environ.get("DOWNLOAD_PATH")
ZIP_PATH = os.environ.get("ZIP_PATH")
SERVICE_ACCOUNT = os.environ.get("SERVICE_ACCOUNT")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")

# --------------------------------------------------------------------------
# Create service with Google API
# --------------------------------------------------------------------------

gc = gspread.service_account(SERVICE_ACCOUNT)
spreadsheet = gc.open_by_key(SPREADSHEET_ID)

CLIENT_SECRET_FILE = SERVICE_ACCOUNT
API_NAME = "sheets"
API_VERSION = "v4"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT, scopes=SCOPES
)

service = build(API_NAME, API_VERSION, credentials=creds)
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
        # vvv Test variable for weekly sheet
        # day = "Friday"
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
        wait.until(EC.element_to_be_clickable((By.ID, "NavBar-Analytics")))
        print("    Grabbing metrics ...\n")
        time.sleep(4)
        analytics = driver.find_element(By.ID, "NavBar-Analytics")
        hover = action.move_to_element(analytics)
        hover.perform()
        time.sleep(1)
        network_analytics = driver.find_element(
            By.XPATH, '//*[@id="app-body"]/div/div[1]/ul/li[7]/div[2]/div/div[2]/div'
        )
        action.move_to_element(network_analytics).click().perform()
        time.sleep(2)
        sle_page = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//*[@id="app-body"]/div/div[2]/div[2]/header/span[1]/div[1]',
                )
            )
        )
        sle_page.click()
        time.sleep(1)
        sle_page = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//*[@title="SLE"]',
                )
            )
        )
        sle_page.click()
        time.sleep(1)
        dropdown = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//*[@id="app-body"]/div/div[2]/div[2]/div[1]/div[1]/span[2]/div/div[2]',
                )
            )
        )
        dropdown.click()
        timeframe = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//*[@id="app-body"]/div/div[2]/div[2]/div[1]/div[1]/span[2]/div/div[2]/div/div/div/div/div[2]/div[1]',
                )
            )
        )
        timeframe.click()
        time.sleep(2)
        download = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//*[@id="app-body"]/div/div[2]/div[2]/div[1]/div[1]/span[2]/div/div[3]/div[2]/button',
                )
            )
        )
        try:
            os.remove(ZIP_PATH)
        except:
            pass
        time.sleep(2)
        download.click()
        # Wait until the file exists
        while not os.path.exists(ZIP_PATH):
            time.sleep(1)
        # Open .zip then convert .csv file to list
        with ZipFile(ZIP_PATH) as zf:
            with zf.open("Sites by Clients.csv", "r") as infile:
                reader = csv.reader(TextIOWrapper(infile, "utf-8"))
                client_data = list(reader)
        try:
            # Delete zip file after
            os.remove(ZIP_PATH)
        except:
            pass
        if day == "Friday":
            dropdown.click()
            time.sleep(1)
            timeframe = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//*[@id="app-body"]/div/div[2]/div[2]/div[1]/div[1]/span[2]/div/div[2]/div/div/div/div/div[1]/div[3]',
                    )
                )
            )
            time.sleep(1)
            timeframe.click()
            time.sleep(2)
            try:
                os.remove(ZIP_PATH)
            except:
                pass
            download.click()
            while not os.path.exists(ZIP_PATH):
                time.sleep(1)
            with ZipFile(ZIP_PATH) as zf:
                with zf.open("SLE List.csv", "r") as infile:
                    reader = pd.read_csv(infile)
                    reader = reader.reset_index()
                    sites = []
                    values = []
                    reader["Site"] = reader["Site"].astype("string")
                    reader["Successful Connect"] = reader["Successful Connect"].astype(
                        "string"
                    )
                    for index, row in reader.iterrows():
                        if len(row["Site"]) <= 3:
                            sites.append([row["Site"]])
                            values.append(re.findall("\d+", row["Successful Connect"]))
                    sle_data = [sites, values]
            try:
                os.remove(ZIP_PATH)
            except:
                pass
            finally:
                driver.quit()
                return [client_data, sle_data]
        else:
            driver.quit()
            return [client_data]
    except Exception as e:
        print(e)
        driver.quit()
        retry()


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
    for i in range(len(data[0])):
        sites = data[0][i][0]
        clients = data[0][i][1]
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
    worksheet_name = "DailyClientCount!"
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

    if len(data) > 1:
        columns = spreadsheet.get_worksheet(2).row_values(1)
        columns_array = []
        title_array = []
        new_column = False
        for el in columns:
            title_array.append(el.split(", "))
            columns_array.append(el.split(", "))
        del columns_array[:2]
        next_row = next_available_row(spreadsheet.get_worksheet(2))
        values = [
            [today],
            ["=average(C%s:J%s)" % (next_row, next_row)],
        ]
        for i in range(len(data[1][0])):
            sites = data[1][0][i]
            connects = data[1][1][i][0]
            # Check if the site from the data file is already in spreadsheet, if it is then match the data with the site
            if sites in columns_array:
                columns_array[columns_array.index(sites)][0] = connects
            # If site does not match any column in spreadsheet then add a new column
            else:
                title_array.append([sites])
                columns_array.extend([[connects]])
                new_column = True
            if sites != "Sites":
                print("Site: " + sites[0] + " | Successful Connects: " + connects + "%")

        # Check if value is a string then erase it (for example if a site has 0 clients it will not show up in the list)
        for list in columns_array:
            for idx in range(len(list)):
                if list[idx] == today or list[idx] == day:
                    continue
                elif list[idx][0].isalpha():
                    list[idx] = " "
        values.extend(columns_array)

        print("\nUpdating Google sheet ...\n")
        worksheet_name = "WeeklySuccessfulConnects!"
        cell_range_insert = "A3"

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


def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list) + 2)


def bye():
    try:
        sys.exit(1)
    except:
        os._exit(1)


def retry():
    sys.stdout.flush()
    os.execv(sys.executable, [sys.executable, '"' + sys.argv[0] + '"'] + sys.argv[1:])


# --------------------------------------------------------------------------
# Function:     main()
# Purpose:      Run the previous functions and catches errors
# --------------------------------------------------------------------------


def main():
    os.system(sys_cls_clear)
    # Don't run on weekends???
    if day == "Saturday" or day == "Sunday":
        driver.quit
        bye()
    else:
        try:
            print("Today is " + day + " " + today + "\n")
            parse_and_upload(scrape())
            print("Done!\n")
        # Retry program if it fails
        except Exception as e:
            print(e)
            try:
                driver.quit()
            except:
                pass
            finally:
                retry()
        finally:
            bye()


main()
