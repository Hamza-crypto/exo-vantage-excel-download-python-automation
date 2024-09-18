from playwright.sync_api import Playwright, sync_playwright
import playwright.async_api
from datetime import datetime, timedelta
import pandas as pd
import time
import os

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
        
with sync_playwright() as playwright:

    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(storage_state="auth.json")
    page = context.new_page()
    page.goto(BASE_URL)
    
    login(page, context)
    time.sleep(2)
    # Initial dropdown selection
    page.get_by_label("Date Type").select_option("Audit Passed")
    page.get_by_text("No", exact=True).click()
    
   
    # Click to open the dropdown so options get loaded
    page.locator("#SchemeId + div .selectize-input").click()

    # Get all dynamically loaded options from the dropdown
    options = page.locator(".selectize-dropdown-content .option")

    # Loop through each option and select it
    option_count = options.count()
    
    print(option_count)
    # page.pause()
    # time.sleep(2)
    
    for i in range(option_count):
        page.locator("#SchemeId + div .selectize-input").click()
        option_value = options.nth(i).get_attribute("data-value")
        option_text = options.nth(i).inner_text()
        
        if option_value:  # Skip empty value
            print(f"Selecting scheme: {option_text} value: {option_value}")
            
            # Click the option to select it
            options.nth(i).click()
            
            # Fill in the date range and click export
            page.locator("#DateFrom").fill('01-Sep-2024')
            page.locator("#DateTo").fill('28-Sep-2024')
            sleep(1)
            
            continue
            
            # Export the report for the current scheme
            page.locator("#ExportButton").click()
            with page.expect_download() as download_info:
                page.locator("#ExportButton").click()
            download = download_info.value
          
            exit()
            # Save the file with a unique name per scheme
            save_file(download, f"{option_text}-EMVIC-VCUSTOMER-INVOICE-SUMMARY-REPORT-1")

            time.sleep(2)  # Wait a bit between downloads
       
    time.sleep(2)
    context.close()
    browser.close()
