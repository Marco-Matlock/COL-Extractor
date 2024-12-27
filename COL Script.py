import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import logging
import re

# Set up logging
logging.basicConfig(
    filename='col_script.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

logging.info("Script started.")

# Google Sheets setup
json_keyfile = r"C:\Users\marco\Desktop\Catalogue of Life\catalogue-of-life-4c5a1e91c118.json"
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]

# Authenticate and connect to Google Sheets
try:
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile, scope)
    client = gspread.authorize(creds)
    logging.info("Successfully authenticated with Google Sheets.")
except Exception as e:
    logging.error(f"Error authenticating with Google Sheets: {e}")
    raise

# Open the Google Sheet
try:
    sheet = client.open("COL Collector").sheet1
    logging.info("Google Sheet opened successfully.")
except Exception as e:
    logging.error(f"Error opening Google Sheet: {e}")
    raise

# Get all plant names from column A, starting from A2
try:
    plant_names = sheet.col_values(1)[1:]  # Skip the header row (A1)
    logging.info(f"Retrieved {len(plant_names)} plant names from the sheet.")
except Exception as e:
    logging.error(f"Error reading plant names from Google Sheets: {e}")
    raise

# Find the first empty cell in column B, starting from B2
try:
    values_in_col_b = sheet.col_values(2)  # Get values in column B
    start_row = len(values_in_col_b) + 1 if '' in values_in_col_b else 2  # Find first empty cell, or start at B2
    logging.info(f"Starting from row {start_row} in column A.")
except Exception as e:
    logging.error(f"Error finding the first empty cell in column B: {e}")
    raise

# Selenium setup for Microsoft Edge
edge_driver_path = r"D:\Python\Drivers\msedgedriver.exe"

# Configure Edge options
edge_options = EdgeOptions()
edge_options.add_argument("--disable-gpu")
edge_options.add_argument("--window-size=1920,1080")
edge_options.add_argument("--headless")

# Initialize EdgeDriver
try:
    service = EdgeService(executable_path=edge_driver_path)
    driver = webdriver.Edge(service=service, options=edge_options)
    logging.info("EdgeDriver initialized successfully.")
except Exception as e:
    logging.error(f"Error initializing EdgeDriver: {e}")
    raise

# Function to extract Taxon ID from URL
def extract_taxon_id(url):
    match = re.search(r'/taxon/([0-9A-Z]+)', url)
    return match.group(1) if match else None

# Set the starting row to # directly
start_row = 1026
logging.info(f"Starting from row {start_row} in column A.")

# Loop through each plant name, starting from row 4543
for row_index, plant_name in enumerate(plant_names[start_row - 2:], start=start_row):
    try:
        # Open the Catalogue of Life website
        driver.get("https://www.catalogueoflife.org/")
        logging.info(f"Opened Catalogue of Life website for plant: {plant_name}.")

        # Wait for the search bar to be visible
        search_box = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "input.ant-input"))
        )

        # Enter the plant name in the search bar
        search_box.clear()
        search_box.send_keys(plant_name)

        # Wait for the dropdown options to load
        dropdown_option = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".ant-select-item.ant-select-item-option"))
        )

        # Add a 1-second delay before clicking the active (first) dropdown option
        time.sleep(1)

        # Click the active (first) dropdown option
        dropdown_option.click()
        logging.info(f"Clicked the active dropdown option for plant: {plant_name}.")

        # Wait for the URL to change and retrieve the current URL
        WebDriverWait(driver, 10).until(EC.url_contains('/taxon/'))
        current_url = driver.current_url

        # Extract the taxon ID from the URL
        if 'taxon' in current_url:
            taxon_id = extract_taxon_id(current_url)
            logging.info(f"Found taxon ID: {taxon_id} for plant: {plant_name}.")
        else:
            taxon_id = None

        # Update the Google Sheet with the taxon ID
        try:
            if taxon_id:
                sheet.update_cell(row_index, 2, taxon_id)  # Update corresponding cell in column B
                logging.info(f"Updated row {row_index} with taxon ID: {taxon_id}.")
            else:
                logging.warning(f"No valid Taxon ID found for plant '{plant_name}' at row {row_index}.")
                sheet.update_cell(row_index, 2, "No ID")
        except Exception as e:
            logging.error(f"Error updating Google Sheet for plant '{plant_name}' at row {row_index}: {e}")
            sheet.update_cell(row_index, 2, "Error")
            continue

        # Add a short delay to avoid getting blocked
        time.sleep(2)

    except Exception as e:
        logging.error(f"Error processing plant '{plant_name}' at row {row_index}: {e}")
        sheet.update_cell(row_index, 2, "Error")
        continue

# Clean up and close the driver
driver.quit()
logging.info("Script completed.")
print("Script completed.")