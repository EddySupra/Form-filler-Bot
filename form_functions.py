from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from pprint import pprint as pp
import gspread
from selenium.common.exceptions import TimeoutException
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import traceback 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import random
import matplotlib.pyplot as plt
from matplotlib.patches import PathPatch
from matplotlib.path import Path
import numpy as np
import re
from PIL import Image
import base64
from pywinauto import Application
import os


# Google Sheets API setup
scope = ["https://spreadsheets.google.com/feeds", 
        "https://www.googleapis.com/auth/spreadsheets", 
        "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('service_account.json', scope)
client = gspread.authorize(creds)
print("Google Sheets access")


# Delay for program actions 2 - 4 secconds 
DELAY = 1 

class PopupCheckException(Exception):
    """Raised when a pop-up is detected during demographic_page processing."""
    pass

def highlight_row_red(sheet, row_number):
    """
    Highlights the specified row in the Google Sheet red.
    :param sheet: Google Sheets worksheet object.
    :param row_number: The row number to highlight (1-indexed).
    """
    try:
        # Define the red background color
        red_color = {"red": 1.0, "green": 0.0, "blue": 0.0}

        # Create a batch update request for row background color
        sheet.format(f"{row_number}:{row_number}", {
            "backgroundColor": red_color
        })

        print(f"Row {row_number} has been highlighted in red.")
    except Exception as e:
        print(f"Error while highlighting row {row_number} in red: {e}")

def highlight_row_green(sheet, row_number):
    """
    Highlights the specified row in the Google Sheet green.
    :param sheet: Google Sheets worksheet object.
    :param row_number: The row number to highlight (1-indexed).
    """
    try:
        # Define the green background color
        green_color = {"red": 0.0, "green": 1.0, "blue": 0.0}

        # Create a batch update request for row background color
        sheet.format(f"{row_number}:{row_number}", {
            "backgroundColor": green_color
        })

        print(f"Row {row_number} has been highlighted in green.")
    except Exception as e:
        print(f"Error while highlighting row {row_number} in green: {e}")

def highlight_row_blue(sheet, row_number):
    """
    Highlights the specified row in the Google Sheet blue.
    :param sheet: Google Sheets worksheet object.
    :param row_number: The row number to highlight (1-indexed).
    """
    try:
        # Define the blue background color
        blue_color = {"red": 0.0, "green": 0.0, "blue": 1.0}

        # Create a batch update request for row background color
        sheet.format(f"{row_number}:{row_number}", {
            "backgroundColor": blue_color
        })

        print(f"Row {row_number} has been highlighted in blue.")
    except Exception as e:
        print(f"Error while highlighting row {row_number} in blue: {e}")

def safe_click(driver, element, timeout=10):
    """
    Clicks an element safely, waiting for Angular modals/overlays to disappear first.
    """
    try:
        # Wait until no modal overlays are present
        WebDriverWait(driver, timeout).until_not(
            EC.presence_of_element_located((By.CSS_SELECTOR, "app-modal, .cdk-overlay-backdrop, .modal-backdrop"))
        )

        # Scroll into view and try normal click
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.3)
        element.click()
    except Exception as e:
        print(f"Standard click failed ({e}); retrying with JS click.")
        try:
            driver.execute_script("arguments[0].click();", element)
        except Exception as js_e:
            print(f"JS click also failed: {js_e}")


def setup_driver(service, options):
    from selenium import webdriver
    return webdriver.Chrome(service=service, options=options)

def speed_test(driver):
    # Wait for the button to be clickable
    ok_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-flat")))
    
    # Click the button
    ok_button.click()
    time.sleep(DELAY)
    print("Speed test screen finshed ")


def login(driver, username_text, password_text):
    try:
        # Locate username and password fields
        username = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "username")))
        password = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "password")))

        # Enter credentials
        username.send_keys(username_text)
        time.sleep(DELAY)
        password.send_keys(password_text)
        time.sleep(DELAY)

        # Click login button
        login_button = driver.find_element("css selector", ".btn.submit-btn.login-btn")
        login_button.click()
        time.sleep(DELAY)
        
        # Wait for page to load after login
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        print("Login successful.")

    except Exception as e:
        print("Error during login:", e)
        raise

def image_to_coordinates(image_path, threshold=128):
    """
    Convert a signature image to a list of (x, y) coordinates.
    Args:
        image_path (str): Path to the image file.
        threshold (int): Pixel intensity threshold to consider a pixel "black".
    Returns:
        List of tuples: Coordinates (x, y) for black pixels.
    """
    img = Image.open(image_path).convert('L')  # Convert image to grayscale
    width, height = img.size
    coordinates = []

    for y in range(height):
        for x in range(width):
            if img.getpixel((x, y)) < threshold:  # Threshold for black pixels
                coordinates.append((x, y))
    
    # Offset coordinates to center the signature on the canvas
    min_x = min(coord[0] for coord in coordinates)
    min_y = min(coord[1] for coord in coordinates)
    centered_coordinates = [(x - min_x, y - min_y) for x, y in coordinates]

    return centered_coordinates

def scale_coordinates(coordinates, canvas_size):
    max_x = max(coord[0] for coord in coordinates)
    max_y = max(coord[1] for coord in coordinates)

    scale_x = canvas_size['width'] / max_x
    scale_y = canvas_size['height'] / max_y
    scale_factor = min(scale_x, scale_y)  # Maintain aspect ratio

    scaled_coordinates = [(int(x * scale_factor), int(y * scale_factor)) for x, y in coordinates]
    return scaled_coordinates

def center_coordinates(coordinates, canvas_size):
    min_x = min(coord[0] for coord in coordinates)
    min_y = min(coord[1] for coord in coordinates)

    offset_x = (canvas_size['width'] // 2) - (max(coord[0] for coord in coordinates) // 2)
    offset_y = (canvas_size['height'] // 2) - (max(coord[1] for coord in coordinates) // 2)

    centered_coordinates = [(x + offset_x, y + offset_y) for x, y in coordinates]
    return centered_coordinates

def normalize_coordinates(coordinates, canvas_size):
    """Normalize coordinates to fit the canvas dimensions."""
    canvas_width, canvas_height = canvas_size['width'], canvas_size['height']
    max_x = max(coord[0] for coord in coordinates)  # Max x in input coordinates
    max_y = max(coord[1] for coord in coordinates)  # Max y in input coordinates
    
    normalized = []
    for x, y in coordinates:
        # Scale x and y to canvas dimensions
        normalized_x = int(x * canvas_width / max_x) if max_x > 0 else 0
        normalized_y = int(y * canvas_height / max_y) if max_y > 0 else 0
        normalized.append((normalized_x, normalized_y))
    return normalized

# Function to generate JavaScript for drawing signature
def generate_signature_js(image_path, canvas_width=512, canvas_height=220):
    with open(image_path, "rb") as img_file:
        base64_str = base64.b64encode(img_file.read()).decode("utf-8")
    js_code = f"""
const canvas = document.querySelector('canvas');
const ctx = canvas.getContext('2d');
const img = new Image();
img.onload = function() {{
    ctx.drawImage(img, 0, 0, {canvas_width}, {canvas_height});
}};
img.src = 'data:image/png;base64,{base64_str}';
"""
    return js_code

def draw_signature(driver):
    try:
        # Wait for the canvas to be present
        canvas = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "canvas"))
        )
        canvas.click()

        # Generate and inject JavaScript
        signature_js = generate_signature_js("C:\\Users\\Work\\Desktop\\sig.png")
        driver.execute_script(signature_js)
        print("Signature added successfully.")
        
        # Wait for the first name element to be present and visible
        first_name_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "firstName")))

        # Clear the input field and type agent's first name
        first_name_input.clear()
        first_name_input.send_keys("Samantha")
        print("Agent's first name entered successfully.")
        time.sleep(DELAY)

        # Wait for the last name element to be present and visible
        last_name_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "lastName")))

        # Clear the input field and type agent's last name
        last_name_input.clear()
        last_name_input.send_keys("Cruz")
        print("Agent's last name entered successfully.")
        time.sleep(DELAY)

        # Wait for the location element to be present and visible
        location_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "location")))

        # Clear the input field and type agent's location
        location_input.clear()
        location_input.send_keys("Upland")
        print("Agent's location entered successfully.")
        time.sleep(DELAY)

        # Wait for the address element to be present and visible
        address_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "address")))

        # Clear the input field and type agent's address
        address_input.clear()
        address_input.send_keys("1931 N Campus Ave")
        print("Agent's address entered successfully.")
        time.sleep(DELAY)

        try:
            # Try to wait for the dropdown element up to 10 seconds
            state_dropdown = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "state"))
            )
            select = Select(state_dropdown)
            select.select_by_value("CA")
            print("California (CA) selected successfully.")
        except TimeoutException:
            print("State dropdown not found on this page. Skipping state selection.")
        except NoSuchElementException:
            print("The 'CA' option was not found in the dropdown.")
        finally:
            time.sleep(1)

        # Click start order
        start_order = driver.find_element(By.CSS_SELECTOR, ".btn.submit-btn.step-btn")
        start_order.click()
        print("Signature page completed")
        time.sleep(DELAY)

    except Exception as e:
        print("Error during signature drawing:", e)
        raise

def eligibility_page(driver):
    try:
        print("print eligibility page started")
        '''
        #BYOD question 
        # Wait for the element to be present and visible
        byod_button = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "label[for='byodNo']")))

        # Check if the BYOD NO button is already selected
        if not byod_button.is_selected():
            # Click the BYOD NO button if it is not selected
            byod_button.click()
            print("BYOD NO button was not selected. It is now clicked.")
            time.sleep(.5)
        else:
            print("BYOD NO is already selected.")
            time.sleep(.5)'''

        # Another CA Enrollment Answer No
        another_ca_input  = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input#AnotherCaEnrollmentAnswerNo")))
        if not another_ca_input.is_selected():
            print("Another CA Enrollment Answer 'No' is not selected. Clicking it now.")
            another_ca_label = driver.find_element(By.CSS_SELECTOR, "label[for='AnotherCaEnrollmentAnswerNo']")
            another_ca_label.click()
            time.sleep(.5)
        else:
            print("Another CA Enrollment Answer 'No' is already selected. Skipping.")
            time.sleep(.5)

        # State Attestation Yes
        state_attestation_input = driver.find_element(By.CSS_SELECTOR, "input#stateAttestationYes")
        if not state_attestation_input.is_selected():
            print("State Attestation 'Yes' is not selected. Clicking it now.")
            state_attestation_label = driver.find_element(By.CSS_SELECTOR, "label[for='stateAttestationYes']")
            state_attestation_label.click()
            time.sleep(.5)
        else:
            print("State Attestation 'Yes' is already selected. Skipping.")
            

        # Freeze Questions Yes
        freeze_questions_input = driver.find_element(By.CSS_SELECTOR, "input#freezeQuestionsYes")
        if not freeze_questions_input.is_selected():
            print("Freeze Questions 'Yes' is not selected. Clicking it now.")
            freeze_questions_label = driver.find_element(By.CSS_SELECTOR, "label[for='freezeQuestionsYes']")
            freeze_questions_label.click()
            time.sleep(.5)
        else:
            print("Freeze Questions 'Yes' is already selected. Skipping.")
            

        # Wait for the Federally-recognized Tribal lands NO to be present and visible
        isTribalNo = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "label[for='isTribalNo']")))
        # Check if the Federally-recognized Tribal lands NO is already selected
        if not isTribalNo.is_selected():
            # Click the radio button if it is not selected
            isTribalNo.click()
            print("Radio button 'isTribalNo' was not selected. It is now clicked.")
            time.sleep(.5)
        else:
            print("Radio button 'isTribalNo' is already selected.")


        # Enrolled Program (CASNAP)
        enrolled_program_input = driver.find_element(By.CSS_SELECTOR, "input#CASNAP")
        if not enrolled_program_input.is_selected():
            print("Enrolled Program 'CASNAP' is not selected. Clicking it now.")
            enrolled_program_label = driver.find_element(By.CSS_SELECTOR, "label[for='CASNAP']")
            enrolled_program_label.click()
            time.sleep(1)
        else:
            print("Enrolled Program 'CASNAP' is already selected. Skipping.")
            

        # Next Button
        next_button = driver.find_element(By.CSS_SELECTOR, "button.btn.submit-btn.step-btn")
        print("Clicking the 'Next' button.")
        next_button.click()
        time.sleep(DELAY)
        print("Eligibility page completed successfully.")

    except Exception as e:
        print("Error during eligibility page processing:", e)
        raise

def demographic_page(driver, row_number,address_row):
    try:
        # Access the Google Sheet
        spreadsheet = client.open("Elsa Leads")  # Replace with your sheet's actual name
        sheet = spreadsheet.sheet1  # Access the first sheet in the file
        row_data = sheet.row_values(row_number)  # Retrieve the row based on row_number
        address_data = sheet.row_values(address_row)  # Address details
        print(f"Using address from row {address_row}")
        print(f"Row {row_number} pulled")

        # Fill address fields using address_data
        address_fields = {
            "address": {"selector": (By.CSS_SELECTOR, '[placeholder="Address"]'), "value": address_data[5]},
            "apt": {"selector": (By.CSS_SELECTOR, '[placeholder="APT/Floor/Other"]'), "value": address_data[6]},
            "city": {"selector": (By.CSS_SELECTOR, '[placeholder="City"]'), "value": address_data[7]},
            "zip": {"selector": (By.CSS_SELECTOR, '[placeholder="Zip"]'), "value": address_data[8]},
        }

        for field_name, field_info in address_fields.items():
            print(f"Filling out field: {field_name}")
            selector_type, selector_value = field_info["selector"]
            field_element =  WebDriverWait(driver, 20).until(EC.presence_of_element_located((selector_type, selector_value)))
            field_element.clear()
            field_element.send_keys(field_info["value"])
            time.sleep(.5)

        print("Address fields populated successfully.")

        # Extract and handle date components
        dob = row_data[2]  # Assuming the date is in column 3 (index 2)
        if '-' in dob:
            month, day, year = dob.split('-')  # Split the date if separated by '-'
        elif '/' in dob:
            month, day, year = dob.split('/')  # Split the date if separated by '/'
        else:
            raise ValueError("Unsupported date format")
        
        adjusted_month = str(int(month) - 1)  # Adjust to 0-based indexing

        # Wait for and select the month
        month_dropdown_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[formcontrolname="month"]')))
        month_dropdown = Select(month_dropdown_element)
        month_dropdown.select_by_value(adjusted_month)
        time.sleep(.5)

        # Wait for and select the day
        day_dropdown_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[formcontrolname="day"]')))
        day_dropdown = Select(day_dropdown_element)
        day_dropdown.select_by_value(day.lstrip("0"))  # Remove leading zeros
        time.sleep(1)

        # Wait for and select the year
        year_dropdown_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[formcontrolname="year"]')))
        year_dropdown = Select(year_dropdown_element)
        year_dropdown.select_by_value(year)
        time.sleep(.5)

        print("Date of Birth selected successfully.")
        
        # Map row data to form fields (update selectors and data mapping as needed)
        form_fields = {
            "first_name": {"selector": (By.CSS_SELECTOR, '[formcontrolname="FirstName"]'), "value": row_data[0]},
            "last_name": {"selector": (By.CSS_SELECTOR, '[formcontrolname="LastName"]'), "value": row_data[1]},
            "email": {"selector": (By.CSS_SELECTOR, '[formcontrolname="Email"]'), "value": row_data[4]},
        }

        # Fill out the form
        for field_name, field_info in form_fields.items():
            print(f"Filling out the field: {field_name}")  # Debugging purpose
            selector_type, selector_value = field_info["selector"]
            field_element = driver.find_element(selector_type, selector_value)
            field_element.send_keys(field_info["value"])
            time.sleep(.5)
        


        print("Handling SSN field")
        ssn_value = row_data[3]  # Assuming the SSN value is in the 4th column (index 3)
        ssn_field = driver.find_element(By.CSS_SELECTOR, '[formcontrolname="Ssn"]')

        # Hide obstructing header
        header = driver.find_element(By.TAG_NAME, 'header')
        driver.execute_script("arguments[0].style.visibility = 'hidden';", header)

        # Scroll the SSN field into view
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", ssn_field)

        # Use JavaScript to click the field
        driver.execute_script("arguments[0].click();", ssn_field)
        # Press left arrow key 4 times
        for _ in range(4):
            ssn_field.send_keys(Keys.ARROW_LEFT)
            time.sleep(.5)  # Optional delay for visibility

        # Input the SSN character by character
        for char in ssn_value:
            ssn_field.send_keys(char)
            time.sleep(.5)
        
        print("SSN field successfully filled.")

        # Restore header visibility
        driver.execute_script("arguments[0].style.visibility = 'visible';", header)
            
        
        # Billing Address
        BillingAddressLabel = driver.find_element(By.CSS_SELECTOR, 'label[for="billingAddressSame"]')
        BillingAddressLabel.click()
        time.sleep(1)
        

        # Permanent Address
        PermanentAddress = driver.find_element(By.CSS_SELECTOR, 'label[for="permanent"]')
        PermanentAddress.click()
        time.sleep(1)

        # Shipping Address
        ShippingAddress = driver.find_element(By.CSS_SELECTOR, 'label[for="shippingAddressSameResidence"]')
        ShippingAddress.click()
        time.sleep(1)

        # Email button
        Email = driver.find_element(By.CSS_SELECTOR, 'label[for="reachQuestionEmail"]')
        Email.click()
        time.sleep(1)

        #Wait for the checkbox to be present and visible
        marketing_checkbox = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "label[for='MarketingConsent']")))
        marketing_checkbox.click()
        time.sleep(1)

        #Next Page 
        next_button = driver.find_element(By.CSS_SELECTOR, "button.btn.submit-btn.step-btn")
        next_button.click()
        time.sleep(1)
        

        # Call popup_check
        if popup_check(driver):
            print("Pop-up detected. Exiting demographic_page.")
            raise PopupCheckException()


        print("Demographic page finshed")
    except PopupCheckException:
        # Pass the exception to the caller to handle it
        raise    
    except Exception as e:
        print("Error during form demographic:", e)
        raise

def popup_check(driver):
    try:
        # Wait for the presence of the pop-up
        popup = WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.modal-dialog")))
        if popup:
            print("Pop-up detected")

            # Extract the text content from the modal body
            popup_text = popup.find_element(By.CSS_SELECTOR, "div.modal-body").text.strip()
            print(f"Pop-up text: {popup_text}")
            
            # Normalize popup text for comparison
            normalized_text = popup_text.lower()

            # Handle the pop-up content
            if "currently this customer cannot apply in person" in normalized_text:
                time.sleep(1)
                driver.back()
                print("Person cannot register at this time. Moving on to next.")
                time.sleep(1)
                return True
            elif "in order to continue this enrollment with entouch," in normalized_text:
                time.sleep(1)
                print("please have the customer contact Enrollment Support .")
                driver.back()
                time.sleep(1)
                return True
            elif "duplicate customer found in entouch wireless" in normalized_text:
                print("duplicate customer found entouch.")
                driver.back()
                time.sleep(1)
                return True
            elif "currently receiving the benefit" in normalized_text:
                print("Someone at the address is already receiving the benefit.")
                # Wait for the button to be clickable
                time.sleep(1)
                ok_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-flat")))

                # Click the button
                time.sleep(1)
                ok_button.click()
                time.sleep(1)
                print("OK button clicked successfully.")
                
            elif "someone at the address provided is currently receiving" in normalized_text:
                print ("Someone at the address is already receiving the benefit")
                # Wait for the button to be clickable
                time.sleep(1)
                ok_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-flat")))

                # Click the button
                time.sleep(1)
                ok_button.click()
                time.sleep(1)
                print("OK button clicked successfully.")
                # Ensure this function exists and is defined.

            elif "the california lifeline administrator has determined you already have service with another lifeline carrier" in normalized_text:
                print("already regisiterd in last 60 days")
                time.sleep(1)
                driver.back()
                time.sleep(1)
                return True
                
            elif "you are already receiving a lifeline discount benefit" in normalized_text:
                print("Transfer is available.")
                driver.back()
                return True
            elif " an error has occurred " in normalized_text:
                print("application error has occured ")
                driver.back()
                return True
            elif "We were unable to confirm the address provided with the united states postal service. if you wish to modify the address, click modify. if the address is correct as entered, to complete your order," in normalized_text:
                print("cant confirm address will ask for proof")
                # Wait for the button to be clickable
                time.sleep(1)
                ok_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-flat")))

                # Click the button
                time.sleep(1)
                ok_button.click()
                time.sleep(1)
                print("OK button clicked successfully.")
                

            elif "the california lifeline administrator has determined you already have an application" in normalized_text:
                print("already a user")
                driver.back()
                return True
            elif "we are unable to enroll you in the federal program at this time" in normalized_text:
                print("already a user")
                driver.back()
                return True
            elif "validation successful" in normalized_text:
                print("validation successful")
                # Wait for the button to be clickable
                ok_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-flat")))

                # Click the button
                ok_button.click()
                time.sleep(1) 
            elif "there was a problem processing your order. Please try again, if the problem persists please contact entouch wireless" in normalized_text:
                print("problem processing order")
                driver.back()
                return True
            elif "we were unable to confirm the address provided with the united states postal service" in normalized_text:
                print("Bad Address, Modify it to good one ")
                # Wait for the OK button to be clickable
                ok_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/simple-modal-holder/simple-modal-wrapper/div/app-modal/div/div/div[3]/button[1]"))
                )

                # Scroll it into view
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", ok_button)
                time.sleep(2)

                # Try normal click first
                try:
                    ok_button.click()
                    print("✅ Clicked 'OK' button normally.")
                except Exception as e:
                    print(f"⚠️ Normal click failed: {e}")
                    try:
                        # Try ActionChains click
                        actions = ActionChains(driver)
                        actions.move_to_element(ok_button).pause(0.2).click().perform()
                        print("✅ Clicked 'OK' button with ActionChains.")
                    except Exception as e2:
                        print(f"⚠️ ActionChains click failed: {e2}")
                        # Final fallback — JavaScript click
                        driver.execute_script("arguments[0].click();", ok_button)
                        print("✅ Clicked 'OK' button with JavaScript.")
                return False
            elif "you can apply for lifeline. you live in a household that does not get lifeline"in normalized_text:
                # Wait until any modals or overlays disappear
                WebDriverWait(driver, timeout).until_not(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "app-modal, .cdk-overlay-backdrop, .modal-backdrop"))
                )

                # Wait for the OK button to be clickable
                ok_button = WebDriverWait(driver, timeout).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-flat[type='button']"))
                )

                # Scroll into view and click safely
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", ok_button)
                time.sleep(3)
                ok_button.click()
                print("✅ OK button clicked successfully.")
            else:
                print("Unhandled pop-up message detected.")
                return False
    except Exception as e:
        print('no pop up continue')


def disclosures_page(driver):
    try:
                # Wait until the <h1> element with the text "Disclosures" appears on the screen
        try:
            # Wait until the <h1> element with class "title" is visible and contains "Disclosures"
            element = WebDriverWait(driver, 0).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "h1.title"))  # Wait for element to appear
            )
            
            # Wait for the text inside the element to be "Disclosures"
            WebDriverWait(driver, 15).until(
                EC.text_to_be_present_in_element((By.CSS_SELECTOR, "h1.title"), "Disclosures")
            )
            print('disclosures page started')
        except Exception as e:
            print(f"Error: {e}")
        

        try:
            # Locate and click 'IehAnotherAdultYes' radio button
            IehAnotherAdult = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'label[for="IehAnotherAdultYes"]'))
            )
            safe_click(driver, IehAnotherAdult)
            print("Clicked the 'IehAnotherAdultYes' radio button successfully.")
        except TimeoutException:
            print("'IehAnotherAdultYes' radio button not found. Continuing...")
 
        try:
            # Locate and click 'IehDiscountNo' radio button
            IehDiscount = WebDriverWait(driver, 1).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'label[for="IehDiscountNo"]'))
            )
            safe_click(driver, IehDiscount)
            print("Clicked the 'IehDiscountNo' radio button successfully.")

            time.sleep(1)

            try:
                ok_button = WebDriverWait(driver, 7).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/simple-modal-holder/simple-modal-wrapper/div/app-modal/div/div/div[3]/button'))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", ok_button)
                time.sleep(0.5)

                actions = ActionChains(driver)
                actions.move_to_element(ok_button).click().perform()
                print("✅ 'Ok' button clicked with ActionChains.")
                time.sleep(1)

            except Exception as e:
                print(f"⚠️ Could not click 'Ok' button with ActionChains: {e}")

        except TimeoutException:
                print("'IehDiscountNo' radio button not found. Continuing...")

        # Try to locate and click the 'CertificationB' checkbox
        try:
            certification_checkbox = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'label[for="CertificationB"]'))
            )
            time.sleep(1)
            safe_click(driver,certification_checkbox)
            print("Clicked the 'CertificationB' checkbox successfully.")
        except TimeoutException:
            print("'CertificationB' checkbox not found. Continuing...")

        # Handle disclosure labels in bulk
        disclosure_selectors = [
            'disclosures3', 'disclosures4', 'disclosures7', 'disclosures8',
            'disclosures9', 'disclosures10', 'disclosures11', 'disclosures12',
            'disclosures13', 'disclosures14', 'disclosures15', 'disclosures16',
            'disclosures17', 'disclosures18', 'disclosures19', 'disclosures21',
            'disclosures22', 'disclosures24', 'disclosures25', 'disclosures26',
            'disclosures27', 'disclosures28', 'disclosures23', 'disclosures29'
        ]

        try:
            # Locate all disclosure labels at once
            labels = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, ', '.join([f'label[for="{d}"]' for d in disclosure_selectors])))
            )

            # Click each label
            for label in labels:
                try:
                    safe_click(driver, label)
                    print(f"Clicked the label successfully: {label.get_attribute('for')}")
                except Exception as e:
                    print(f"Error clicking label '{label.get_attribute('for')}': {e}")

        except TimeoutException:
            print("Some or all disclosure labels not found.")

                # Wait for the canvas to be present
        canvas = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "canvas"))
        )
        safe_click(driver, canvas)

        # Generate and inject JavaScript
        signature_js = generate_signature_js("C:\\Users\\Work\\Desktop\\sig.png")
        driver.execute_script(signature_js)
        print("Signature added successfully.")

        # Try to locate and click the 'I Accept' button
        try:
            accept_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.submit-btn.step-btn'))
            )
            safe_click(driver, accept_button)
            
            print("Clicked the 'I Accept' button successfully.")
            time.sleep(1)
        except TimeoutException:
            print("'I Accept' button not found. Continuing...")

        print ("Disclosure page done")
        time.sleep(1)
        
    except Exception as e:
        print("Error during elegiblity", e)
        raise




def wait_for_user():
    print("Waiting... Type '1' to proceed.")
    while True:
        user_input = input("Type '1' to continue: ").strip().lower()
        if user_input == "1":
            break  # Exit loop when "good" is entered
        


def upload_file(driver, file_path):
    try:
        time.sleep(1)
        upload_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/div[1]/div/section/app-image-capture/div/form/div/div[2]/div/div/div/div/label")))
        time.sleep(1)
        upload_button.click()
        print("Upload button clicked.")
        
        time.sleep(2)  # Wait for the file dialog to appear

        # Step 2: Use Pywinauto to handle the file dialog
        app = Application().connect(title_re="Open")  # Matches window title (varies by OS)
        file_dialog = app.window(title_re="Open")

        # Step 3: Enter the file path
        file_dialog.Edit.type_keys(file_path, with_spaces=True)

        # Step 4: Click the "Open" button
        file_dialog.Open.click()
        print("File uploaded successfully.")

    except Exception as e:
        print("Error uploading file:", e)

def get_random_jpg(folder_path):
    """
    Picks a random .jpg or .jpeg file from the given folder.

    :param folder_path: Path to the folder containing image files.
    :return: Full path of a randomly selected .jpg or .jpeg file.
    """
    try:
        # Include both .jpg and .jpeg extensions
        image_files = [f for f in os.listdir(folder_path)
                       if f.lower().endswith(('.jpg', '.jpeg'))]
        if not image_files:
            raise FileNotFoundError("No .jpg or .jpeg files found in the folder.")
        return os.path.join(folder_path, random.choice(image_files))
    except Exception as e:
        print(f"Error selecting random image: {e}")
        return None


def service_type_check(driver, file_path):
    try:
        try:
            # Wait until the <h1> element with class "title" is visible and contains "Image Capture"
            element = WebDriverWait(driver, 0).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "h1.title"))
            )
            
            # Wait for the text inside the element to be "Image Capture"
            WebDriverWait(driver, 0).until(
                EC.text_to_be_present_in_element((By.CSS_SELECTOR, "h1.title"), "Image Capture")
            )
            
            print("Element is fully loaded and text is 'Image Capture'.")
    
        except Exception as e:
            print(f"Error: {e}")

        # -------------------------
        # Check for "Proof of Identification"
        # -------------------------
        try:
            proof_id_element = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Proof of Identification')]"))
            )
            if proof_id_element and "Proof of Identification" in proof_id_element.text:
                print("Proof of Identification, cannot run app")
                for i in range(3):
                    driver.back()
                    print(f"Navigated back {i+1} time(s).")
                    time.sleep(1)
                return True
        except:
            pass

        # -------------------------
        # Check for "CalFresh"
        # -------------------------
        try:
            proof_calfresh_element = WebDriverWait(driver, 0.5).until(
                EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'CalFresh')]"))
            )
            if proof_calfresh_element and "CalFresh" in proof_calfresh_element.text:
                print("Proof of CALFRESH, cannot run app")
                for i in range(3):
                    driver.back()
                    print(f"Navigated back {i+1} time(s).")
                    time.sleep(1)
                return True
        except:
            pass

        # -------------------------
        # Check for "Address Proof"
        # -------------------------
        proof_address_present = False
        try:
            proof_address_element = WebDriverWait(driver, 0.5).until(
                EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Address Proof')]"))
            )
            if proof_address_element and "Address Proof" in proof_address_element.text:
                print("Proof of Address detected.")
                proof_address_present = True
        except:
            pass

        # -------------------------
        # Check for "Agent Selfie"
        # -------------------------
        try:
            agent_selfie_present = False
            agent_selfie_element = WebDriverWait(driver, 0.5).until(
                EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Agent Selfie')]"))
            )
            if agent_selfie_element and "Agent Selfie" in agent_selfie_element.text:
                agent_selfie_present = True
        except:
            pass

        # -------------------------
        # If Agent Selfie is present (alone or with Address Proof)
        # -------------------------
        if proof_address_present and agent_selfie_present:
            print("Both 'Proof of Address' and 'Agent Selfie' detected — treating as Agent Selfie Only.")
        if agent_selfie_present:
            print("Agent Selfie Only, good app")
            time.sleep(1)
            print("Start code for image input")
            upload_file(driver, file_path)
            print("Please click submit button")
            return False

        # -------------------------
        # If only Address Proof without Agent Selfie
        # -------------------------
        if proof_address_present and not agent_selfie_present:
            print("Proof of Address only — change address to a good one")
            wait_for_user()

    except Exception as e:
        print("Service type page check error or element not found:", e)


def finish_process(driver, esn_row, imei_row):
    try:
        print("Input IMEI")
        
        # Access the Google Sheet
        try:
            # Define the scope for accessing Google Sheets and Drive
            scope = [
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]

            # Load credentials from the JSON key file
            creds = ServiceAccountCredentials.from_json_keyfile_name('service_account.json', scope)

            # Authorize the client
            client = gspread.authorize(creds)
            print("Google Sheets access authorized successfully.")

            inventory_spreadsheet = client.open("Sam Inv")  # Replace with your sheet's actual name
            inv_sheet = inventory_spreadsheet.sheet1  # Access the first sheet in the file
            esn_row_inv = inv_sheet.row_values(esn_row)  # Retrieve the esn row based on esn_row_number
           

            print(f"Row {esn_row} pulled: {esn_row_inv[0]}")
        except Exception as e:
            print(f"Error accessing Google Sheet or fetching row {esn_row}: {e}")
            return
        esn_not_valid = True
        while esn_not_valid == True:
            esn_row +=1
            esn_row_inv = inv_sheet.row_values(esn_row)  # Retrieve the esn row based on imei_row_number
            print(f"Row {esn_row} pulled: {esn_row_inv[1]}")
            # Locate the 'CurrentEsn' input box
            try:
                textbox = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='CurrentEsn']"))
                )
                textbox.click()
                textbox.send_keys(esn_row_inv[0])
                print("Entered IMEI into 'CurrentEsn' input box.")
                time.sleep(1)
            except Exception as e:
                print(f"Error locating or interacting with 'CurrentEsn': {e}")
                return

            # Validate IMEI
            try:
                validate_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-flat.validate-btn"))
                )
                validate_button.click()
                print("Clicked the 'Validate' button.")
                time.sleep(1)
            except Exception as e:
                print(f"Error clicking 'Validate' button: {e}")
                return

            # Handle 'OK' button
            try:
                time.sleep(1)    
                # Wait for the modal to appear
                modal = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "modal-dialog"))
                )
                # Find the element containing the text
                modal_text = modal.find_element(By.CSS_SELECTOR, "div.modal-body > p").text
                time.sleep(1)    
                
                if "not found" in modal_text:
                    print("esn not found. Taking appropriate action.")
                    # Example: Click the OK button
                    ok_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-flat[type="button"]')))
                    time.sleep(1)      
                    ok_button.click()
                    time.sleep(1)  
                    esn_not_valid = True   
                else:
                    print("Modal text does not contain 'not found'.")
                    # Example: Click the OK button
                    ok_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-flat[type="button"]')))      
                    ok_button.click()
                    esn_not_valid = False
            except Exception as e:
                print("Error handling 'not found' modal:", e)

        # Handle 'CurrentImei' input
        imei_not_valid = True
        while imei_not_valid == True:
            imei_row +=1
            imei_row_inv = inv_sheet.row_values(imei_row)  # Retrieve the esn row based on imei_row_number
            print(f"Row {imei_row} pulled: {imei_row_inv[1]}")

            try:
                current_imei_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="CurrentImei"]'))
                )
                current_imei_input.click()
                current_imei_input.clear()
                time.sleep(1)
                current_imei_input.send_keys(imei_row_inv[1])
                print("Entered IMEI into 'CurrentImei' input field.")
                time.sleep(1)
            except Exception as e:
                print(f"Error interacting with 'CurrentImei': {e}")
                return
            
            try:
                # Wait for the button to become enabled
                validate_button1 = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/app-root/div[1]/div/section/app-service-type/form/div/div[2]/div[1]/div[3]/button"))
                )
                WebDriverWait(driver, 20).until(
                    lambda driver: validate_button1.get_attribute("disabled") is None
                )

                # Force-click using JavaScript
                driver.execute_script("arguments[0].click();", validate_button1)
                print("Clicked the 'Validate' button using JavaScript.")
            except TimeoutException:
                print("'Validate' button not clickable or not found.")
            try:
                time.sleep(1)    
                # Wait for the modal to appear
                modal = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "modal-dialog"))
                )
                # Find the element containing the text
                modal_text = modal.find_element(By.CSS_SELECTOR, "div.modal-body > p").text
                time.sleep(1)    
                
                if "not found" in modal_text:
                    print("IMEI not found. Taking appropriate action.")
                    # Example: Click the OK button
                    ok_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-flat[type="button"]')))
                    time.sleep(1)      
                    ok_button.click()
                    time.sleep(1)
                    imei_not_valid = True   
                else:
                    print("Modal text does not contain 'not found'.")
                    # Example: Click the OK button
                    ok_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-flat[type="button"]')))      
                    ok_button.click()
                    imei_not_valid = False
            except Exception as e:
                print("Error handling 'not found' modal:", e)

        # Handle PIN input
        try:
            pin_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="Pin"]'))
            )
            pin_input.send_keys("1234")
            print("Entered PIN into 'Pin' input field.")
            time.sleep(1)
        except Exception as e:
            print(f"Error interacting with 'Pin': {e}")
            return

        # Handle Confirm PIN input
        try:
            confirm_pin_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="ConfirmPin"]'))
            )
            confirm_pin_input.send_keys("1234")
            print("Entered PIN into 'ConfirmPin' input field.")
            time.sleep(1)
        except Exception as e:
            print(f"Error interacting with 'ConfirmPin': {e}")
            return

        # Handle SelectedCaIncomeRange
        try:
            income_range_radio = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'label[for="incomeRanges11"]')))
            
            income_range_radio.click()
            print("Clicked the 'SelectedCaIncomeRange' radio button.")
            time.sleep(1)
        except Exception as e:
            print(f"Error interacting with 'SelectedCaIncomeRange': {e}")
            return

        # Handle SelectedCaGender
        try:
            gender_radio = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'label[for="californiaGenders4"]'))
            )
            gender_radio.click()
            print("Clicked the 'SelectedCaGender' radio button.")
            time.sleep(1)
        except Exception as e:
            print(f"Error interacting with 'SelectedCaGender': {e}")
            return

        # Handle SelectedCaRace
        try:
            race_radio = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'label[for="californiaRace7"]'))
            )
            race_radio.click()
            print("Clicked the 'SelectedCaRace' radio button.")
            time.sleep(1)
        except Exception as e:
            print(f"Error interacting with 'SelectedCaRace': {e}")
            return

        # Handle Submit Order
        try:
            submit_order_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.submit-btn.step-btn'))
            )
            submit_order_button.click()
            print("Clicked the 'Submit Order' button successfully.")
            time.sleep(1)
        except Exception as e:
            print(f"Error clicking 'Submit Order' button: {e}")
            return
        print ("order complete wait 10-15 min and new app will start")
        
        highlight_row_green(inv_sheet,esn_row)
        
        #timer inbetween apps
        time.sleep(random.randint(600,900))
        
        # Locate the 'New Order' link using its attributes
        new_order_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.btn.btn-flat[href="/agentsignature"]'))
        )
        # Click the 'New Order' link
        new_order_link.click()
        print("Clicked the 'New Order' link successfully.")
        

    except Exception as e:
        print(f"General error in finish_process: {e}")


