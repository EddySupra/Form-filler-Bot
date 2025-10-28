import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from form_functions import *
import time
import random

def setup_driver_with_geolocation(latitude, longitude, accuracy=100):
    """
    Set up an undetected Chrome WebDriver with geolocation override.
    """
    options = uc.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--start-maximized")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36")

    # Optional: set Chrome binary location (if needed)
    options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"

    driver = uc.Chrome(options=options)

    # Set geolocation
    driver.execute_cdp_cmd("Emulation.setGeolocationOverride", {
        "latitude": latitude,
        "longitude": longitude,
        "accuracy": accuracy
    })
    print(f"Geolocation set to Latitude: {latitude}, Longitude: {longitude}, Accuracy: {accuracy}")
    return driver


# Prompt user for geolocation input
try:
    latitude = float(input("Enter the latitude: "))
    longitude = float(input("Enter the longitude: "))
except ValueError:
    print("Invalid input. Please enter numerical values for latitude and longitude.")
    exit(1)

# --- build chrome options ---
chrome_options = Options()

# 1) point to your actual Chrome install:
chrome_options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"

# Setup driver
driver = setup_driver_with_geolocation(latitude, longitude, accuracy=100)
tracker = 0 #complete application tracker

#file_path = "C:\\Users\\casas\\Desktop\\IMG_0472.jpg"

folder_path = "C:\\Users\\Work\\Desktop\\IMAGES"  # Change to your actual folder path


try:
    wait_for_user()
    wait_for_user()

    # Open the application page
    driver.get("https://app.cgmllc.net/")
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    print("Application page loaded.")
    
    
    # Bypass speed test screen
    speed_test(driver)

    # Log in to the application7854
    login(driver, "V1_Samanthacruz1", "Water1233!!")

    row_number = 1  # Start with the first row
    address_row = 1  # Start with the first row for address details
    spreadsheet = client.open("Sam Leads")  # Replace with your sheet's actual name
    sheet = spreadsheet.sheet1  # Access the first sheet in the file
    esn_row = 1
    imei_row = 1 
    
    while tracker != 20:
        try:
            # Outer loop starts with drawing the signature
            print("Drawing the signature...")
            draw_signature(driver)

            restart_eligibility = True
            while restart_eligibility:
                try:
                    file_path = get_random_jpg(folder_path)

                    # Increment row_number for each new person
                    row_number += 1
                    address_row +=1

                    print("Filling out eligibility page...")
                    eligibility_page(driver)

                    print("Filling out demographic page...")
                    print(f"Processing row {row_number} with address from row {address_row}")
                    demographic_page(driver, row_number,address_row)
                    time.sleep(1)

                    # Check for pop-up
                    if popup_check(driver):
                        print('Cannot register this person, starting eligibility again.')
                        highlight_row_red(sheet, row_number)
                        restart_eligibility = True  # Explicitly restart eligibility loop
                        continue  # Restart eligibility process

                    print("Filling out disclosures page...")
                    disclosures_page(driver)

                    # Check service type
                    if service_type_check(driver, file_path):
                        print("Service type issue, restarting from eligibility.")
                        highlight_row_red(sheet, row_number)
                        continue  # Restart eligibility process
                    
                    # Move past eligibility if successful
                    restart_eligibility = False
                
                except Exception as e:
                    print(f"An error occurred during eligibility steps: {e}")
                    highlight_row_red(sheet, row_number)
                    continue  # Restart eligibility process

          

            # Example usage
            print("Pausing execution...")
            time.sleep(1)
            wait_for_user()
            time.sleep(1)
            print("Resuming script...")

            # Input IMEI info
            print("Processing IMEI...")
            finish_process(driver, esn_row, imei_row)

            # Highlight the row green on success
            highlight_row_green(sheet, row_number)
            esn_row += 1
            imei_row += 1
            tracker += 1
            address_row += 1
            print(f"Tracker updated: {tracker}")
        
        except PopupCheckException:
            print("Pop-up detected during demographic_page. Restarting eligibility loop.")
            highlight_row_red(sheet, row_number)
            continue  # Restart eligibility process
        
        except Exception as e:
            print(f"An error occurred: {e}")
        
    print("20 apps finished.")
    driver.quit()

except Exception as e:
    print("BIG error occurred:", e)
finally:
    driver.quit()
