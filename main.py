from playwright.sync_api import Playwright, sync_playwright
import playwright.async_api
from datetime import datetime, timedelta
import pandas as pd
import time
import os



BASE_URL = "https://ecovantage.alitsy.com/Finance/CertificateBilling"

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
        page.get_by_placeholder("E-mail").fill(username)
        page.get_by_placeholder("Password").fill(password)
        page.get_by_role("button", name="Log In").click()
        context.storage_state(path="auth.json")
        
        
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

    current_date1 = datetime.now()
    one_year_ago = (current_date1 - timedelta(days=365)).strftime("%d-%m-%Y")
    current_date2 = datetime.now()
    next_day = (current_date2 + timedelta(days=1)).strftime("%d-%m-%Y")

    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(storage_state="auth.json")
    page = context.new_page()
    page.goto(BASE_URL)
    
    login(page, context)
    
    # Initial dropdown selection
    page.get_by_label("Date Type").select_option("Audit Passed")
    page.get_by_text("No", exact=True).click()
    
    page.get_by_text("ESS Lighting").click()
    
    page.locator("#DateFrom").fill('01-Sep-2024')
    page.locator("#DateTo").fill('29-Sep-2024')
    
    page.locator("#ExportButton").click()
    with page.expect_download() as download_info:
        page.locator("#ExportButton").click()
    download = download_info.value
    save_file(download, "EMVIC-VCUSTOMER-INVOICE-SUMMARY-REPORT-1")
       
    time.sleep(2)
    context.close()
    browser.close()
