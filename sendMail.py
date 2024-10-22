import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

import time
import pyperclip  # For copying HTML to clipboard
from email_content import generate_html_email_content

CHROME_WINDOW_SIZE = "1920,1080"
DOWNLOAD_DIR = r"/path/to/your/folder"
PAGE_LOAD_WAIT_SEC = 60
DELAY_SEC = 3

with open('authen.json') as json_file:
    authen = json.load(json_file)

def sign_in():


    chrome_options = Options()
    chrome_options.add_argument(f"--window-size={CHROME_WINDOW_SIZE}")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
    })

    driver = webdriver.Chrome(options=chrome_options)
    driver_wait = WebDriverWait(driver, PAGE_LOAD_WAIT_SEC)

    # Navigate to outlook.office.com/mail
    driver.get("https://outlook.office.com/mail/")

    time.sleep(DELAY_SEC)

    # Input my email address
    login_box = (By.NAME, "loginfmt")
    login_box = driver_wait.until(EC.presence_of_element_located(login_box))
    login_box.send_keys(authen["username"])
    login_box.send_keys(Keys.RETURN)

    time.sleep(DELAY_SEC)

    # Input my password
    password_box = (By.NAME, "passwd")
    password_box = driver_wait.until(EC.presence_of_element_located(password_box))
    password_box.send_keys(authen["password"])
    password_box.send_keys(Keys.ENTER)

    # Confirm my stay signed-in option
    # Confirm my stay signed-in option
    stay_sign_in = (By.XPATH, "//button[@id='acceptButton']")
    stay_sign_in = driver_wait.until(EC.presence_of_element_located(stay_sign_in))
    stay_sign_in.click()

    time.sleep(DELAY_SEC)

    # # Access the mailbox
    # mail_box_panel = (By.CLASS_NAME, preprocess_element_class("C2IG3 if6B2 oTkSL iDEcr OPUpK"))
    # mail_box_panel = driver_wait.until(EC.element_to_be_clickable(mail_box_panel))
    # mail_box_panel.click()

    # time.sleep(DELAY_SEC)

    # mail_box = (By.CSS_SELECTOR, f"div[title^='{MAIL_BOX}']")
    # mail_box = driver_wait.until(EC.presence_of_element_located(mail_box))
    # mail_box.click()

    # time.sleep(DELAY_SEC)
    
    #popup blocking
    # options = Options()
    # options.add_argument("--disable-popup-blocking")
    # options.add_experimental_option("excludeSwitches", ["enable-automation"])
    # options.add_experimental_option('useAutomationExtension', False)
    # # Initialize WebDriver
    # driver = webdriver.Chrome(options=options)
    # driver.get("https://outlook.office.com/mail/")
    # # Wait and close pop-up by finding its frame or unique element (assumed)
    # WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe.popUpFrame")))
    # driver.find_element(By.CSS_SELECTOR, "button.closePopUp").click()
    # # Switch back to the main content after closing the pop-up
    # driver.switch_to.default_content()


def send_email(driver, to_email, bcc_email, subject, email_body_html):
    """
    Function to send an email with HTML content.
    """
    # Wait for the page to load fully
    time.sleep(5)  # Adjust if necessary, or use WebDriverWait for more control

    # Click the 'New Message' button (compose button)
    #new_message_button = driver.find_element(By.XPATH, "//span[@id='id__177']")
    # new_message_button = driver.find_element(By.XPATH, '//button[@aria-label="New mail"]')
    # new_message_button.click()

    # time.sleep(2)  # Wait for the new message window to open

    # Click "New message" button
    # Access the mailbox
    mail_box_panel = (By.CLASS_NAME, preprocess_element_class("C2IG3 if6B2 oTkSL iDEcr OPUpK"))
    mail_box_panel = driver_wait.until(EC.element_to_be_clickable(mail_box_panel))
    mail_box_panel.click()

    time.sleep(DELAY_SEC)

    mail_box = (By.CSS_SELECTOR, f"div[title^='{MAIL_BOX}']")
    mail_box = driver_wait.until(EC.presence_of_element_located(mail_box))
    mail_box.click()

    time.sleep(DELAY_SEC)
    
    print("Clicking 'New mail' button...")
    try:
        new_message_btn = driver_wait.until(EC.presence_of_element_located((By.XPATH, '//button[@aria-label="New mail"]')))
        new_message_btn.click()
    except Exception as e:
        print(f"Failed to click 'New mail' button: {e}")
        return
    time.sleep(2)
    # Enter the recipient email address
    to_field = driver.find_element(By.XPATH, "//input[@aria-label='To']")
    to_field.send_keys(to_email)
    to_field.send_keys(Keys.RETURN)

    # Show BCC field (if necessary)
    bcc_toggle_button = driver.find_element(By.XPATH, "//button[@aria-label='Show Bcc']")
    bcc_toggle_button.click()

    time.sleep(3)  # Wait for the BCC field to appear

    # Enter BCC email address
    bcc_field = driver.find_element(By.XPATH, "//input[@aria-label='Bcc']")
    bcc_field.send_keys(bcc_email)
    bcc_field.send_keys(Keys.RETURN)

    # Enter the email subject
    subject_field = driver.find_element(By.XPATH, "//input[@aria-label='Add a subject']")
    subject_field.send_keys(subject)

    # Switch to the email body field (this field accepts HTML, but you need to insert the content as plain text)
    body_field = driver.find_element(By.XPATH, "//div[@aria-label='Message body']")

    # Paste the HTML content directly (Outlook accepts HTML in the body)
    body_field.send_keys(Keys.CONTROL, 'v')  # Simulate Ctrl + V (paste)

    # Use pyperclip to copy HTML content to the clipboard
    pyperclip.copy(email_body_html)

    # Wait for the email to be composed properly
    time.sleep(1)

    # Send the email
    send_button = driver.find_element(By.XPATH, "//button[@aria-label='Send']")
    send_button.click()

    print("Email sent successfully with HTML content.")

# Example usage:
if __name__ == "__main__":
    # Assuming you already have a logged-in driver instance
    driver = webdriver.Chrome()  # Or any other driver you are using
    
    # Login function should have already been called at this point
    sign_in()
    
    

    # Replace these with actual values
    recipient_email = "neela.rahman9916@gmail.com"
    bcc_email = "bacbonsohada@outlook.com"
    email_subject = "Test HTML Email Subject"

    # Generate dynamic HTML content for each recipient
    recipient_name = "John Doe"  # You can change this dynamically for each recipient
    email_body_html = generate_html_email_content(recipient_name)

    # Call the send_email function
    send_email(driver, recipient_email, bcc_email, email_subject, email_body_html)

    # You can close the browser after sending the email
    time.sleep(5)  # Optional, to observe the result
    driver.quit()
