import json
from playwright.sync_api import sync_playwright
import time
import os
import subprocess
import pyautogui
import pyperclip
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Define constants for Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = r'C:\Users\dinev\OneDrive\Desktop\Automation Scripts\glassy-tube-427214-q0-76ba5e6df01a.json'
SPREADSHEET_ID = '1DeexQnb_jGpO16krhdkyANP-J9sI_Hqy7NHufQ3ZjvU'
RANGE_NAME = 'Sheet1!A:O'  # Ensure this matches your actual sheet name

def get_sheets_service():
    print("Initializing Google Sheets API service...")
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    return service

def update_google_sheet(order_id, parcel_number):
    print("Updating Google Sheet...")
    service = get_sheets_service()
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    print(f"Updating order ID {order_id} with parcel number {parcel_number}")

    for idx, row in enumerate(values):
        if row and row[0] == order_id:
            cell_O = f'Sheet1!O{idx + 1}'
            cell_P = f'Sheet1!P{idx + 1}'
            print(f"Updating cell {cell_O} with parcel number {parcel_number} and cell {cell_P} with 'Shipped'")
            try:
                update_result_O = sheet.values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=cell_O,
                    valueInputOption='USER_ENTERED',
                    body={'values': [[parcel_number]]}
                ).execute()

                update_result_P = sheet.values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=cell_P,
                    valueInputOption='USER_ENTERED',
                    body={'values': [["Shipped"]]}
                ).execute()

                print(f"Update result for cell {cell_O}: {update_result_O}")
                print(f"Update result for cell {cell_P}: {update_result_P}")
                print(f"Updated cell {cell_O} with parcel number {parcel_number} and cell {cell_P} with 'Shipped'")

                # Add a delay before verification
                time.sleep(5)

                # Verify update
                verify_result_O = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=cell_O).execute()
                verify_result_P = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=cell_P).execute()
                if (verify_result_O.get('values', [[None]])[0][0] == parcel_number and
                    verify_result_P.get('values', [[None]])[0][0] == "Shipped"):
                    print(f"Verification successful: {cell_O} contains {parcel_number} and {cell_P} contains 'Shipped'")
                else:
                    print(f"Verification failed: {cell_O} or {cell_P} does not contain the expected values")

            except Exception as e:
                print(f"Failed to update cell {cell_O} or {cell_P}: {e}")
            break
    else:
        print(f"Order ID {order_id} not found in the sheet")

def open_woocommerce_orders_dashboard():
    print("Opening WooCommerce Orders Dashboard...")
    with sync_playwright() as p:
        # Define the download directory
        download_dir = "C:\\Users\\dinev\\Downloads"
        os.makedirs(download_dir, exist_ok=True)

        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)

        # Load cookies from the file
        cookies_path = "cookies.json"
        if os.path.exists(cookies_path):
            with open(cookies_path, 'r') as cookies_file:
                cookies = json.load(cookies_file)
                context.add_cookies(cookies)
                print("Loaded cookies from file.")

        # Open new page
        page = context.new_page()

        # Go directly to Woocommerce orders dashboard
        page.goto("https://orhidea.bg/wp-admin/admin.php?page=wc-orders")

        # Wait for the page to fully load
        page.wait_for_load_state('networkidle')

        # Select the search box and wait for user input
        search_box_selector = 'input[name="s"]'
        page.click(search_box_selector)
        print("Please type the Order ID in the search box and press Enter in the browser.")

        # Listen for the Enter key press event
        page.evaluate("""
            document.querySelector('input[name="s"]').addEventListener('keypress', function (e) {
                if (e.key === 'Enter') {
                    window.orderID = e.target.value;
                }
            });
        """)

        # Poll for the order ID to be set in the window object
        while True:
            order_id = page.evaluate("window.orderID")
            if order_id:
                break

        print(f"Captured Order ID: {order_id}")

        # Navigate to the specific order page using the captured order ID
        page.goto(f"https://orhidea.bg/wp-admin/post.php?post={order_id}&action=edit")

        # Wait for navigation to the order page
        page.wait_for_load_state('networkidle')

        # Scroll to the bottom of the page
        page.evaluate("window.scrollTo(0, document.body.scrollHeight);")

        # Click the "Create Parcel" button
        create_parcel_button_selector = '#order_create_parcel'  # Updated selector
        page.click(create_parcel_button_selector)

        # Wait for any resulting action to complete
        page.wait_for_load_state('networkidle')

        # Copy the parcel number to the clipboard
        parcel_number_selector = '//td[contains(text(), "Parcel number")]/following-sibling::td'
        parcel_number = page.inner_text(parcel_number_selector)
        pyperclip.copy(parcel_number)
        print(f"Copied Parcel Number: {parcel_number} to clipboard")

        # Wait until the "Parcel print A4" link is fully loaded
        parcel_print_link_selector = '//td[contains(text(), "Parcel print A4:")]/following-sibling::td/a'
        
        # Wait for the link to be available
        page.wait_for_selector(parcel_print_link_selector, state='visible')

        # Wait for 3 seconds before clicking the "Parcel print A4" link
        time.sleep(3)

        # Start the download
        with page.expect_download() as download_info:
            page.click(parcel_print_link_selector)
        
        download = download_info.value
        download_path = os.path.join(download_dir, download.suggested_filename)
        download.save_as(download_path)
        print(f"Downloaded file to: {download_path}")

        # Wait for 5 seconds to ensure the download is complete
        time.sleep(5)

        # Open the downloaded file with Adobe Acrobat
        acrobat_path = "C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe"
        subprocess.Popen([acrobat_path, "/p", download_path])

        # Wait for 5 seconds to ensure Adobe Acrobat has fully loaded the print dialog
        time.sleep(5)

        # Simulate pressing "Enter" to print the document
        pyautogui.press('enter')

        # Wait for 5 seconds to ensure the print command is processed
        time.sleep(5)

        # Close Adobe Acrobat after printing (optional)
        subprocess.run(["taskkill", "/IM", "Acrobat.exe", "/F"], check=True)

        # Ensure the file is deleted after printing
        for _ in range(5):  # Retry up to 5 times
            try:
                os.remove(download_path)
                print(f"Deleted file: {download_path}")
                break
            except Exception as e:
                print(f"Failed to delete file: {download_path}, retrying... {e}")
                time.sleep(2)  # Wait before retrying

        # Update Google Sheet with the parcel number
        update_google_sheet(order_id, parcel_number)

if __name__ == "__main__":
    print("Starting script...")
    open_woocommerce_orders_dashboard()
