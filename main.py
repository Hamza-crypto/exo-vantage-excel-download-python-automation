from playwright.sync_api import Playwright, sync_playwright 
import pandas as pd
import time
import os
import requests
import calendar
from datetime import datetime

BASE_URL = "https://ecovantage.alitsy.com/Finance/CertificateBilling"

# Read configuration from file
def load_config(file_path="config.txt"):
    config = {}
    with open(file_path, "r") as data:
        for line in data:
            key, value = line.strip().split(' = ')
            config[key.lower()] = value if key.lower() != 'number_of_months_to_go_back' else int(value)
    return config

config = load_config()

username = config.get('username')
password = config.get('password')
destination_path = config.get('destination_path')
number_of_months = config.get('number_of_months_to_go_back')
rcti_file = config.get('rctifile', 'false').lower() == 'true'
compliance_file = config.get('compliancefile', 'false').lower() == 'true'
install_product_details_file = config.get('installproductdetailsfile', 'false').lower() == 'true'

def login(page, context):
    time.sleep(1)
    if page.title() == 'Log in':
        print('Logging in...')
        page.get_by_label("E-mail").fill(username)
        page.get_by_label("Password").fill(password)
        page.get_by_role("button", name="Log In").click()
        context.storage_state(path="auth.json")
    else:
        print('Already Logged In')

def save_file(content, filename, extension='csv'):
    downloaded_file_name = f"downloaded.{extension}"
    with open(downloaded_file_name, 'wb') as f:
        f.write(content)

    if extension == 'csv':
        downloaded_file = pd.read_csv(downloaded_file_name)
    else:
        downloaded_file = pd.read_excel(downloaded_file_name)

    current_date_time = datetime.now().strftime("%Y%m%d%H%M%S")
    file_name = f"{destination_path}/{filename}-{current_date_time}.csv"
    downloaded_file.to_csv(file_name, sep='|', index=False)
    print(f"File {file_name} downloaded successfully")
    print()
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
    for i in range(options.count()):
        if int(options.nth(i).get_attribute("data-value")) == option_value:
            options.nth(i).click()
            break

def post_request(session, url, payload, filename, extension='csv'):
    try:
        response = session.post(url, data=payload)
        response.raise_for_status()
        print(f"Downloading the file {filename}")
        save_file(response.content, filename, extension)
    except requests.exceptions.RequestException as e:
        print(f"Request error: {e}")

def process_scheme_options(page, session, scheme_options, date_ranges, report_type="rcti"):
    for option_value, option_name in scheme_options:
        try:
            print(option_name)
            if report_type == 'rcti':
                select_scheme_option(page, option_value)
            for start_date, end_date in date_ranges:
                print()
                print(start_date, end_date)
                if report_type == "rcti":
                    payload = {
                        'submitAction': 'Export',
                        'SchemeId': option_value,
                        'TypeOfDateFilter': 'Audit Passed',
                        'DateFrom': start_date,
                        'DateTo': end_date,
                        'ShowFinalised': 'true'
                    }
                    post_request(session, BASE_URL, payload, f'alitsycertificatebillingcsv', 'xlsx')
                elif report_type == "compliance":
                    payload = {
                        'submitAction': 'ExportCsv',
                        'SchemeId': option_value,
                        'TypeOfDateFilter': 'Audit Assigned',
                        'DateFrom': start_date,
                        'DateTo': end_date,
                        'ExcludeRegistered': 'false'
                    }
                    post_request(session, 'https://ecovantage.alitsy.com/Report/ComplianceSummary', payload, f'alitsycompliancesummarycsv-{option_name.replace(" ", "-")}')
                elif report_type == "install_product_details":
                    payload = {
                        'submitAction': 'ExportCsv',
                        'SchemeList': option_value,
                        'TypeOfDateFilter': 'Commencement Date',
                        'DateFrom': start_date,
                        'DateTo': end_date,
                        'IncludeCancelledJobs': 'true',
                        'IncludeAddOns': 'true'
                    }
                    post_request(session, 'https://ecovantage.alitsy.com/Report/InstallProductDetail', payload, f'alitsyinstallproductdetail-{option_name.replace(" ", "-")}')
        except Exception as e:
            print(f"Error processing scheme {option_name}: {str(e)}")
            print('**************************************************')

def setup_session_and_context(playwright):
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(storage_state="auth.json")
    page = context.new_page()
    page.goto(BASE_URL, timeout=0)
    login(page, context)
    time.sleep(2)
    
    cookies = context.cookies()
    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'])
    
    return session, page, context, browser

with sync_playwright() as playwright:
    date_ranges = get_month_date_ranges(number_of_months)
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

    session, page, context, browser = setup_session_and_context(playwright)
    
    if rcti_file:
        try:
            page.get_by_label("Date Type").select_option("Audit Passed")
            page.get_by_text("No", exact=True).click()
            process_scheme_options(page, session, scheme_options, date_ranges, report_type="rcti")
            print('------------------------RCTI Files Completed--------------------------------')
        except Exception as e:
            print(f"Error processing RCTI files: {str(e)}")
            
    if compliance_file:
        try:
            BASE_URL = 'https://ecovantage.alitsy.com/Report/ComplianceSummary'
            page.goto(BASE_URL, wait_until="networkidle")
            page.get_by_label("Date Type").select_option("Audit Assigned")
            page.get_by_text("Yes", exact=True).nth(1).click()
            process_scheme_options(page, session, scheme_options, date_ranges, report_type="compliance")
            print('------------------------Compliance Files Completed--------------------------------')
        except Exception as e:
            print(f"Error processing Compliance files: {str(e)}")
            
    if install_product_details_file:
        try:
            scheme_options.append((51, "Warranty"))
            BASE_URL = 'https://ecovantage.alitsy.com/Report/InstallProductDetail'
            page.goto(BASE_URL, wait_until="networkidle")
            page.get_by_label("Date Type").select_option("Commencement Date")
            page.get_by_text("No", exact=True).first.click()
            page.get_by_text("No", exact=True).nth(1).click()
            process_scheme_options(page, session, scheme_options, date_ranges, report_type="install_product_details")
            print('------------------------Install Product Details Files Completed--------------------------------')
        except Exception as e:
            print(f"Error processing Install Product Details files: {str(e)}")
    
    print('All files downloaded successfully.')
    print('Browser will autoclose in 10 seconds.')
    time.sleep(10)
    context.close()
    browser.close()

# --load-storage=auth.json