from playwright.sync_api import Playwright, sync_playwright
import playwright.async_api
from datetime import datetime, timedelta
import pandas as pd
import time
import os
import requests

BASE_URL = "https://ecovantage.alitsy.com/Finance/CertificateBilling"
# BASE_URL = "http://localhost:8523/"

data = open("config.txt", "r")
for x in data:
    if 'username' in x:
        username = x.replace('username = ', '').replace('\n', '')
    if 'password' in x:
        password = x.replace('password = ', '').replace('\n', '')
    if 'destination_path' in x:
        destination_path = x.replace('destination_path = ', '').replace('\n', '')  
    if 'number_of_months_to_go_back' in x:
            number_of_months = x.replace('number_of_months_to_go_back = ', '').replace('\n', '')  


def login(page, context):
    time.sleep(1)
    if page.title() == 'Log in':
        print('Logging in ...')
        page.get_by_label("E-mail").fill(username)
        page.get_by_label("Password").fill(password)
        page.get_by_role("button", name="Log In").click()
        context.storage_state(path="auth.json")
    else:
        print('Already Logged In')    
      
        
def save_file(download, file_name):
    
    downloaded_file_name = "downloaded.csv"
    download.save_as(downloaded_file_name)
    
    downloaded_file = pd.read_csv('downloaded.csv')
    current_date_today = datetime.now()
    current_date_today = current_date_today.strftime("%Y%m%d%H%M%S")
    file_name = destination_path + '/' +file_name + current_date_today + '.csv' 
    downloaded_file.to_csv(file_name, sep='|', index=False)
    print(f"File {file_name} downloaded successfully")
    os.remove(downloaded_file_name)    
    time.sleep(2)    


def select_scheme_option(page, option_value):
    # Click the selectize dropdown to load the options
    page.locator("#SchemeId + div .selectize-input").click()

    # Wait for the options to be visible
    options = page.locator(".selectize-dropdown-content .option")
    page.wait_for_selector(".selectize-dropdown-content .option")

    # Loop through each option and select the matching one
    option_count = options.count()
    
    for i in range(option_count):
        current_value = int(options.nth(i).get_attribute("data-value"))
        option_text = options.nth(i).inner_text()
        
        # Match the desired value
        if current_value == option_value:
            print(f"Selecting scheme: {option_text} with value: {current_value}")
            
            # Click on the option to select it
            options.nth(i).click()
            break

# Now use the cookies to simulate the button click and perform a POST request
def post_request_with_saved_session(cookies, scheme_value, date_from, date_to):

    # Prepare the payload for the POST request (same as captured from browser dev tools)
    payload = {
        'submitAction': 'Export',
        'InstallId': '',
        'ProjectId': '',
        'SchemeId': scheme_value,  # Change as necessary
        'AgentId': '',
        'ProjectDescription': '',
        'TypeOfDateFilter': 'Audit Passed',
        'DateFrom': date_from,
        'DateTo': date_to,
        'RctiId': '',
        'ShowFinalised': 'true',
        'ShowFinalised': 'true'
    }

    # Construct headers and include the cookies
    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'])

    # Perform the POST request with the saved session
    response = session.post(BASE_URL, data=payload)

    if response.status_code == 200:
        print("POST request successful, downloading the file...")
        # Save the downloaded file
        file_name = 'downloaded_report.csv'  # Adjust as needed
        with open(file_name, 'wb') as f:
            f.write(response.content)
        print(f"File saved as {file_name}")
    else:
        print(f"Failed to download file. Status Code: {response.status_code}")


with sync_playwright() as playwright:

    scheme_options = [
        (1, "ESS Lighting"),
        (43, "HEER HP"),
        (40, "HEER HVAC"),
        (3, "HEER Lighting"),
        (39, "IHEAB HP"),
        (41, "IHEAB HVAC"),
        (33, "IHEAB RDC"),
        (5, "REPS CL"),
        (37, "REPS HC2A"),
        (44, "REPS HC2B"),
        (42, "REPS HC3"),
        (34, "REPS HP"),
        (35, "REPS RDC1"),
        (21, "SOLAR"),
        (45, "STC HP - DO NOT USE"),
        (32, "VEU APP"),
        (31, "VEU HP"),
        (4, "VEU RESI"),
        (2, "VEU S34"),
        (6, "VEU S35"),
        (38, "VEU S44"),
        (36, "VEU SH")
    ]
    
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(storage_state="auth.json")
    page = context.new_page()
    page.goto(BASE_URL)
    
    login(page, context)
    time.sleep(2)

    # Initial dropdown selection
    # page.get_by_label("Date Type").select_option("Audit Passed")
    # page.get_by_text("No", exact=True).click()
    # select_scheme_option(page, 39)

    date_from = "01-Sep-2024"    
    date_to = "30-Sep-2024"    
    # Get the session cookies after logging in
    cookies = context.cookies()
    post_request_with_saved_session(cookies, 39, date_from, date_to)
    # with page.expect_download() as download_info:
    #     page.frame_locator("iframe >> nth=1").frame_locator("#mainContent").get_by_role("button", name="Export To CSV").click()
    # download = download_info.value
    # save_file(download, "EMVIC-VCUSTOMER-INVOICE-SUMMARY-REPORT-1")

    # for value, text in scheme_options:
    #     select_scheme_option(page, value)
    #     time.sleep(2)
    time.sleep(5)
    page.pause()
    
       
    time.sleep(2)
    context.close()
    browser.close()
