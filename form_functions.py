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



# Google Sheets API setup
scope = ["https://spreadsheets.google.com/feeds", 
        "https://www.googleapis.com/auth/spreadsheets", 
        "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('service_account.json', scope)
client = gspread.authorize(creds)
print("Google Sheets access")


# Delay for program actions 2 - 4 secconds 
DELAY = random.randint(2, 4)

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
        signature_js = generate_signature_js("C:\\Users\\casas\\Desktop\\1111.png")
        driver.execute_script(signature_js)
        print("Signature added successfully.")
        
        # Wait for the first name element to be present and visible
        first_name_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "firstName")))

        # Clear the input field and type agent's first name
        first_name_input.clear()
        first_name_input.send_keys("shijuana")
        print("Agent's first name entered successfully.")
        time.sleep(DELAY)

        # Wait for the last name element to be present and visible
        last_name_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "lastName")))

        # Clear the input field and type agent's last name
        last_name_input.clear()
        last_name_input.send_keys("tinoco")
        print("Agent's last name entered successfully.")
        time.sleep(DELAY)

        # Wait for the location element to be present and visible
        location_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "location")))

        # Clear the input field and type agent's location
        location_input.clear()
        location_input.send_keys("Bell")
        print("Agent's location entered successfully.")
        time.sleep(DELAY)

        # Wait for the address element to be present and visible
        address_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "address")))

        # Clear the input field and type agent's address
        address_input.clear()
        address_input.send_keys("6801 Atlantic Ave")
        print("Agent's address entered successfully.")
        time.sleep(DELAY)

        # Wait for the dropdown element to be present and visible
        state_dropdown = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "state")))

        # Create a Select object from the dropdown element
        select = Select(state_dropdown)

        # Select the option with value "CA"
        select.select_by_value("CA")
        print("California (CA) selected successfully.")
        time.sleep(DELAY)

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
        #BYOD question 
        # Wait for the element to be present and visible
        byod_button = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "label[for='byodNo']")))

        # Check if the BYOD NO button is already selected
        if not byod_button.is_selected():
            # Click the BYOD NO button if it is not selected
            byod_button.click()
            print("BYOD NO button was not selected. It is now clicked.")
            time.sleep(DELAY)
        else:
            print("BYOD NO is already selected.")
            time.sleep(DELAY)

        # Another CA Enrollment Answer No
        another_ca_input = driver.find_element(By.CSS_SELECTOR, "input#AnotherCaEnrollmentAnswerNo")
        if not another_ca_input.is_selected():
            print("Another CA Enrollment Answer 'No' is not selected. Clicking it now.")
            another_ca_label = driver.find_element(By.CSS_SELECTOR, "label[for='AnotherCaEnrollmentAnswerNo']")
            another_ca_label.click()
            time.sleep(DELAY)
        else:
            print("Another CA Enrollment Answer 'No' is already selected. Skipping.")
            time.sleep(DELAY)

        # State Attestation Yes
        state_attestation_input = driver.find_element(By.CSS_SELECTOR, "input#stateAttestationYes")
        if not state_attestation_input.is_selected():
            print("State Attestation 'Yes' is not selected. Clicking it now.")
            state_attestation_label = driver.find_element(By.CSS_SELECTOR, "label[for='stateAttestationYes']")
            state_attestation_label.click()
            time.sleep(DELAY)
        else:
            print("State Attestation 'Yes' is already selected. Skipping.")
            

        # Freeze Questions Yes
        freeze_questions_input = driver.find_element(By.CSS_SELECTOR, "input#freezeQuestionsYes")
        if not freeze_questions_input.is_selected():
            print("Freeze Questions 'Yes' is not selected. Clicking it now.")
            freeze_questions_label = driver.find_element(By.CSS_SELECTOR, "label[for='freezeQuestionsYes']")
            freeze_questions_label.click()
            time.sleep(DELAY)
        else:
            print("Freeze Questions 'Yes' is already selected. Skipping.")
            

        # Wait for the Federally-recognized Tribal lands NO to be present and visible
        isTribalNo = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "label[for='isTribalNo']")))
        # Check if the Federally-recognized Tribal lands NO is already selected
        if not isTribalNo.is_selected():
            # Click the radio button if it is not selected
            isTribalNo.click()
            print("Radio button 'isTribalNo' was not selected. It is now clicked.")
            time.sleep(DELAY)
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
        spreadsheet = client.open("App AI lead")  # Replace with your sheet's actual name
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
            field_element = driver.find_element(selector_type, selector_value)
            field_element.clear()
            field_element.send_keys(field_info["value"])
            time.sleep(1)

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
        time.sleep(1)

        # Wait for and select the day
        day_dropdown_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[formcontrolname="day"]')))
        day_dropdown = Select(day_dropdown_element)
        day_dropdown.select_by_value(day.lstrip("0"))  # Remove leading zeros
        time.sleep(1)

        # Wait for and select the year
        year_dropdown_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[formcontrolname="year"]')))
        year_dropdown = Select(year_dropdown_element)
        year_dropdown.select_by_value(year)
        time.sleep(1)

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
            time.sleep(1)
        


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
            time.sleep(0.1)  # Optional delay for visibility

        # Input the SSN character by character
        for char in ssn_value:
            ssn_field.send_keys(char)
            time.sleep(1)
        
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


        #validation button
        validation_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-flat')))
        validation_button.click()
        time.sleep(1)


        print("Demographic page finshed")
    except PopupCheckException:
        # Pass the exception to the caller to handle it
        raise    
    except Exception as e:
        print("Error during form submission:", e)
        raise

def popup_check(driver):
    try:
        # Wait for the presence of the pop-up
        popup = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.modal-dialog")))
        if popup:
            print("Pop-up detected")

            # Extract the text content from the modal body
            popup_text = popup.find_element(By.CSS_SELECTOR, "div.modal-body").text.strip()
            print(f"Pop-up text: {popup_text}")
            
            # Normalize popup text for comparison
            normalized_text = popup_text.lower()

            # Handle the pop-up content
            if "currently this customer cannot apply in person" in normalized_text:
                driver.back()
                print("Person cannot register at this time. Moving on to next.")
                time.sleep(2)

                return True
            elif "in order to continue this enrollment with entouch" in normalized_text:
                print("please have the customer contact Enrollment Support .")
                driver.back()
                time.sleep(2)
                return True
            elif "duplicate customer found in entouch wireless" in normalized_text:
                print("duplicate customer found entouch.")
                driver.back()
                time.sleep(2)
                return True
            elif "currently receiving the benefit" in normalized_text:
                print("Someone at the address is already receiving the benefit.")
                # Wait for the button to be clickable
                ok_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-flat")))

                # Click the button
                ok_button.click()
                print("OK button clicked successfully.")
                popup_check(driver)  # Ensure this function exists and is defined.
            elif "someone at the address provided is currently receiving" in normalized_text:
                print ("Someone at the address is already receiving the benefit")
                return False

            elif "you are already receiving a lifeline discount benefit" in normalized_text:
                print("Transfer is available.")
                # Wait for the "Yes" button to be clickable
                yes_button = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[text()='Yes' and @class='btn btn-flat' and @type='button']"))
                )
                # Click the button
                yes_button.click()
                print("Yes button clicked successfully.")
                return False
            else:
                print("Unhandled pop-up message detected.")
                return False
    except Exception as e:
        print('no pop up continue')


def disclosures_page(driver):
    try:
        print('disclosures page started')

        # Try to locate and click 'IehAnotherAdultYes'
        try:
            IehAnotherAdult = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'label[for="IehAnotherAdultYes"]'))
            )
            IehAnotherAdult.click()
            print("Clicked the 'IehAnotherAdultYes' radio button successfully.")
            time.sleep(1)
        except TimeoutException:
            print("'IehAnotherAdultYes' radio button not found. Continuing...")

        # Try to locate and click 'IehDiscountNo'
        try:
            IehDiscount = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'label[for="IehDiscountNo"]'))
            )
            IehDiscount.click()
            print("Clicked the 'IehDiscountNo' radio button successfully.")
        except TimeoutException:
            print("'IehDiscountNo' radio button not found. Continuing...")

        # Try to locate and click the 'Ok' button
        try:
            ok_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'button.btn.btn-flat[type="button"]'))
            )
            ok_button.click()
            print("Clicked the 'Ok' button successfully.")
        except TimeoutException:
            print("'Ok' button not found. Continuing...")

        # Try to locate and click the 'CertificationB' checkbox
        try:
            certification_checkbox = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'label[for="CertificationB"]'))
            )
            certification_checkbox.click()
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
            'disclosures27', 'disclosures28'
        ]

        try:
            # Locate all disclosure labels at once
            labels = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, ', '.join([f'label[for="{d}"]' for d in disclosure_selectors])))
            )

            # Click each label
            for label in labels:
                try:
                    label.click()
                    print(f"Clicked the label successfully: {label.get_attribute('for')}")
                except Exception as e:
                    print(f"Error clicking label '{label.get_attribute('for')}': {e}")

        except TimeoutException:
            print("Some or all disclosure labels not found.")

        # Try to locate and click the 'I Accept' button
        try:
            accept_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.submit-btn.step-btn'))
            )
            accept_button.click()
            print("Clicked the 'I Accept' button successfully.")
            time.sleep(1)
        except TimeoutException:
            print("'I Accept' button not found. Continuing...")
        print ("Disclosure page done")
        time.sleep(DELAY)
        
    except Exception as e:
        print("Error during elegiblity", e)
        raise
def service_type_check(driver):
    try:
        print("Checking service type page...")
        try:
            # Check if the element exists
            current_esn = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="CurrentEsn"]')))
            print("Instant Approval")
            return  # Exit the function if the element is found
        except TimeoutException:
            print("not instant approval")


        heading = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h2.title")))
        if heading:
            print("Required CAL FRESH PICTURE found. Restarting the application process.")

            # Navigate back once
            driver.back()
            print("Navigated back once.")

            # Pause briefly to ensure the page loads
            time.sleep(2)

            # Navigate back a second time
            driver.back()
            print("Navigated back twice.")

            time.sleep(2)

            # Navigate back a third time
            driver.back()
            print("Navigated back three times.")

            return True  # Indicate the element was found
        else:

            print("instant approval, continue")
            return False
    except Exception as e:
        print("Service type page check error or element not found.")

def finish_process(driver, row_number_inv):
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

            inventory_spreadsheet = client.open("b1i")  # Replace with your sheet's actual name
            inv_sheet = inventory_spreadsheet.sheet1  # Access the first sheet in the file
            row_inv = inv_sheet.row_values(row_number_inv)  # Retrieve the row based on row_number
            print(f"Row {row_number_inv} pulled: {row_inv}")
        except Exception as e:
            print(f"Error accessing Google Sheet or fetching row {row_number_inv}: {e}")
            return

        # Locate the 'CurrentEsn' input box
        try:
            textbox = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='CurrentEsn']"))
            )
            textbox.click()
            textbox.send_keys(row_inv[0])
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
            ok_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-flat[type="button"]'))
            )
            ok_button.click()
            print("Clicked the 'OK' button.")
            time.sleep(1)
        except Exception as e:
            print(f"Error clicking 'OK' button: {e}")
            return

        # Handle 'CurrentImei' input
        try:
            current_imei_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="CurrentImei"]'))
            )
            current_imei_input.send_keys(row_inv[1])
            print("Entered IMEI into 'CurrentImei' input field.")
            time.sleep(3)
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
            # Locate the 'OK' button
            ok_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-flat[type="button"]'))
            )
            # Click the 'OK' button
            ok_button.click()
            print("Clicked the 'OK' button successfully.")
            time.sleep(1)
        except TimeoutException:
            print("'OK' button not found. Continuing...")

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
        print ("order complete wait 7-8 min and new app will start")
        
        highlight_row_green(inv_sheet,row_number_inv)
        
        #timer inbetween apps
        time.sleep(random.randint(420,480))
        
        # Locate the 'New Order' link using its attributes
        new_order_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.btn.btn-flat[href="/agentsignature"]'))
        )
        # Click the 'New Order' link
        new_order_link.click()
        print("Clicked the 'New Order' link successfully.")
        

    except Exception as e:
        print(f"General error in finish_process: {e}")