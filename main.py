from playwright.sync_api import Playwright, sync_playwright
import pandas as pd
import time
import os
import requests
import calendar
from datetime import datetime

BASE_URL = "https://ecovantage.alitsy.com/Finance/CertificateBilling"

# Read configuration
data = open("config.txt", "r")
for x in data:
    if 'username' in x:
        username = x.replace('username = ', '').replace('\n', '')
    if 'password' in x:
        password = x.replace('password = ', '').replace('\n', '')
    if 'destination_path' in x:
        destination_path = x.replace('destination_path = ', '').replace('\n', '')  
    if 'number_of_months_to_go_back' in x:
        number_of_months = int(x.replace('number_of_months_to_go_back = ', '').replace('\n', ''))  

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

def save_file(content):
    downloaded_file_name = "downloaded.xlsx"
    with open(downloaded_file_name, 'wb') as f:
        f.write(content)

    try:
        downloaded_file = pd.read_excel(downloaded_file_name)
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return

    current_date_time = datetime.now().strftime("%Y%m%d%H%M%S")
    file_name = f"{destination_path}/alitsycertificatebillingcsv-{current_date_time}.csv"
    downloaded_file.to_csv(file_name, sep='|', index=False)
    print(f"File {file_name} downloaded successfully")
    os.remove(downloaded_file_name)
    time.sleep(2)

def get_month_date_ranges(months_to_go_back):
    today = datetime.now()
    date_ranges = []
    
    for i in range(months_to_go_back):
        year = today.year
        month = today.month - i
        
        if month <= 0:
            month += 12
            year -= 1

        first_day = datetime(year, month, 1).strftime("%d-%b-%Y")
        last_day = datetime(year, month, calendar.monthrange(year, month)[1]).strftime("%d-%b-%Y")

        date_ranges.insert(0, (first_day, last_day))
    
    return date_ranges

def select_scheme_option(page, option_value):
    page.locator("#SchemeId + div .selectize-input").click()
    page.wait_for_selector(".selectize-dropdown-content .option")
    options = page.locator(".selectize-dropdown-content .option")
    option_count = options.count()
    
    for i in range(option_count):
        current_value = int(options.nth(i).get_attribute("data-value"))
        option_text = options.nth(i).inner_text()
        
        if current_value == option_value:
            options.nth(i).click()
            break

def post_request_with_saved_session(session, scheme_value, date_from, date_to):
    print(scheme_value, date_from, date_to)
    payload = {
        'submitAction': 'Export',
        'InstallId': '',
        'ProjectId': '',
        'SchemeId': scheme_value,
        'AgentId': '',
        'ProjectDescription': '',
        'TypeOfDateFilter': 'Audit Passed',
        'DateFrom': date_from,
        'DateTo': date_to,
        'RctiId': '',
        'ShowFinalised': 'true'
    }
    
    try:
        response = session.post(BASE_URL, data=payload)
        response.raise_for_status()
        print("Downloading the file...")
        save_file(response.content)    
    except requests.exceptions.RequestException as e:
        print(f"Request error: {e}")

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
    
    date_ranges = get_month_date_ranges(number_of_months)
    
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(storage_state="auth.json")
    page = context.new_page()
    page.goto(BASE_URL, wait_until="networkidle")
    
    login(page, context)
    time.sleep(2)
    
    cookies = context.cookies()
    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'])
    
    page.get_by_label("Date Type").select_option("Audit Passed")
    page.get_by_text("No", exact=True).click()

    for option in scheme_options:
        option_value = int(option[0])
        select_scheme_option(page, option_value)
        for start_date, end_date in date_ranges:
            post_request_with_saved_session(session, option_value, start_date, end_date)
        print('--------------------------------------------------------')
    
    time.sleep(2)
    context.close()
    browser.close()
