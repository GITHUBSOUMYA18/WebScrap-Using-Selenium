import sys
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import re

# Suppress unwanted debug and error messages
sys.stderr = open(os.devnull, 'w')  # Suppress errors in stderr

# URLs of the Microsoft update pages
urls = [
    "https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date",
    "https://learn.microsoft.com/en-us/officeupdates/update-history-current-channel-preview",
    "https://learn.microsoft.com/en-us/officeupdates/update-history-beta-channel"
]

# User input: Build version (Example: "18025.20140")
build_version = input("Enter the build version: ").strip()
build_pattern = rf"\(Build {re.escape(build_version)}\)"  # Matches (Build 18025.20140)

# Set up Selenium WebDriver
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--headless")  # Run in headless mode to avoid opening a browser window
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

found_results = []

# Function to get header name by extracting the last three parts of the URL
def extract_header_name(url):
    url_parts = url.split('/')
    header_parts = url_parts[-3:]  # Extract last three parts of the URL
    
    # Clean and format the header name
    clean_header = ' '.join(header_parts[1:]).replace('-', ' ').title()  # Skip 'en-us' and clean the header
    return clean_header

# Loop through each website
for url in urls:
    driver.get(url)
    wait = WebDriverWait(driver, 15)  # Wait for elements to appear

    # Ensure full page loads
    time.sleep(5)

    # Scroll down multiple times to load all data (lazy loading)
    for _ in range(5):
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.PAGE_DOWN)
        time.sleep(1)

    try:
        # Step 1: Try finding the build version inside a table
        tables = driver.find_elements(By.TAG_NAME, "table")
        found_in_table = False

        for table in tables:
            rows = table.find_elements(By.TAG_NAME, "tr")

            # Extract column headers (first row of the table)
            headers = [header.text.strip() for header in rows[0].find_elements(By.TAG_NAME, "th")]

            for row in rows[1:]:  # Skip the first row since it contains headers
                cells = row.find_elements(By.TAG_NAME, "td")
                row_data = [cell.text.strip() for cell in cells]

                # Check if build version exists in any column
                if any(build_version in cell for cell in row_data):
                    found_in_table = True

                    # Get the index of the column where the build version was found
                    col_index = next(i for i, cell in enumerate(row_data) if build_version in cell)

                    # Extract "Version XXXX" from the same row
                    version_info = next((cell for cell in row_data if "Version" in cell), "Version Not Found")

                    # Clean output for found result
                    print(f"✅ Found in table at {url}")
                    print(f"📌 Column Name: {headers[col_index]}")
                    print(f"📄 Extracted Version: {version_info}")

                    found_results.append((headers[col_index], version_info))

        # Step 2: If not found in table, check normal text
        if not found_in_table:
            elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Build')]")

            for element in elements:
                text_content = element.text.strip()
                match = re.search(build_pattern, text_content)

                if match:
                    # Extract the header name from the last three parts of the URL
                    header_name = extract_header_name(url)

                    # Clean output for found result
                    print(f"✅ Found outside table at {url}: {text_content}")
                    print(f"📌 Column Name: {header_name}")  # Dynamically generated header name
                    print(f"📄 Extracted Version: {text_content}")

                    found_results.append((header_name, text_content))

    except Exception as e:
        continue

# Close the browser
driver.quit()

# Print final message if no results were found
if not found_results:
    print("\n❌ Build version is not available in the website.")
