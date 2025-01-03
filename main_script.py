from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from form_functions import *
import time

def setup_driver_with_geolocation(service, options, latitude, longitude, accuracy=100):
    """
    Set up the WebDriver with geolocation override.
    """
    driver = webdriver.Chrome(service=service, options=options)

    # Set geolocation
    params = {
        "latitude": latitude,
        "longitude": longitude,
        "accuracy": accuracy
    }
    driver.execute_cdp_cmd("Emulation.setGeolocationOverride", params)
    print(f"Geolocation set to Latitude: {latitude}, Longitude: {longitude}, Accuracy: {accuracy}")
    return driver

# Prompt user for geolocation input
try:
    latitude = float(input("Enter the latitude: "))
    longitude = float(input("Enter the longitude: "))
except ValueError:
    print("Invalid input. Please enter numerical values for latitude and longitude.")
    exit(1)

# Set up ChromeDriver service
chrome_service = Service("C:\\Users\\casas\\Desktop\\chromedriver-win64\\chromedriver.exe")
chrome_options = Options()
chrome_options.add_argument("--incognito")

# Initialize the driver with geolocation
driver = setup_driver_with_geolocation(chrome_service, chrome_options, latitude, longitude)
tracker = 0 #complete application tracker

try:
    # Open the application page
    driver.get("https://app.cgmllc.net/")
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    print("Application page loaded.")
    
    # Bypass speed test screen
    speed_test(driver)

    # Log in to the application
    login(driver, "V1_ShijuanaTinoco1", "Water123!!")

    row_number = 1  # Start with the first row
    address_row = 2  # Start with the first row for address details
    spreadsheet = client.open("App AI lead")  # Replace with your sheet's actual name
    sheet = spreadsheet.sheet1  # Access the first sheet in the file
    row_number_inv = 1
    
    while tracker != 13:
        try:
            # Outer loop starts with drawing the signature
            print("Drawing the signature...")
            draw_signature(driver)

            restart_eligibility = True
            while restart_eligibility:
                try:
                    # Increment row_number for each new person
                    row_number += 1
                    print("Filling out eligibility page...")
                    eligibility_page(driver)

                    print("Filling out demographic page...")
                    print(f"Processing row {row_number} with address from row {address_row}")
                    demographic_page(driver, row_number,address_row)

                    # Check for pop-up
                    if popup_check(driver):
                        print('Cannot register this person, starting eligibility again.')
                        highlight_row_red(sheet, row_number)
                        continue  # Restart eligibility process

                    print("Filling out disclosures page...")
                    disclosures_page(driver)

                    # Check service type
                    if service_type_check(driver):
                        print("Service type issue, restarting from eligibility.")
                        highlight_row_red(sheet, row_number)
                        continue  # Restart eligibility process
                    
                    # Move past eligibility if successful
                    restart_eligibility = False
                
                except Exception as e:
                    print(f"An error occurred during eligibility steps: {e}")
                    highlight_row_red(sheet, row_number)
                    continue  # Restart eligibility process

            # Input IMEI info
            print("Processing IMEI...")
            finish_process(driver, row_number_inv)

            # Highlight the row green on success
            highlight_row_green(sheet, row_number)
            row_number_inv += 1
            tracker += 1
            address_row += 1
            print(f"Tracker updated: {tracker}")
        
        except PopupCheckException:
            print("Pop-up detected during demographic_page. Restarting eligibility loop.")
            highlight_row_red(sheet, row_number)
            continue  # Restart eligibility process
        
        except Exception as e:
            print(f"An error occurred: {e}")
        
    print("16 apps finished.")
    driver.quit()

except Exception as e:
    print("BIG error occurred:", e)
finally:
    driver.quit()