from playwright.sync_api import Playwright, sync_playwright
import playwright.async_api
from datetime import datetime, timedelta
import pandas as pd
import time
import os
import requests
import calendar

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

    # Try reading the temporary file with the correct encoding
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
    
    # Generate date ranges for each month starting from the current month
    for i in range(months_to_go_back):
        # Calculate the year and month (i=0 will give current month)
        year = today.year
        month = today.month - i
        
        # Adjust for year change when going back from January
        if month <= 0:
            month += 12
            year -= 1

        # Get the first and last day of the month
        first_day = datetime(year, month, 1).strftime("%d-%b-%Y")
        last_day = datetime(year, month, calendar.monthrange(year, month)[1]).strftime("%d-%b-%Y")

        # Append to the list
        date_ranges.insert(0, (first_day, last_day))
    
    return date_ranges


# Now use the cookies to simulate the button click and perform a POST request
def post_request_with_saved_session(session, scheme_value, date_from, date_to):

    print(scheme_value, date_from, date_to)
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
    
    try:
        response = session.post(BASE_URL, data=payload)
        if response.status_code == 200:
            print("Downloading the file...")
            # Save the downloaded file
            save_file(response.content)    
        else:
            print(f"Failed to download file. Status Code: {response.status_code}")
    except e:
        print(f"Error {e}")

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
    # Construct headers and include the cookies
    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'])
    
    count = 1
    for option in scheme_options:
        option_value = int(option[0])
        print(option[1])
        for start_date, end_date in date_ranges:
            
            post_request_with_saved_session(session, option_value, start_date, end_date)
        print('--------------------------------------------------------')
        count = count + 1
    
       
    time.sleep(2)
    context.close()
    browser.close()
