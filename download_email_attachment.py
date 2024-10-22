import json
import re
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


CHROME_WINDOW_SIZE = "1920,1080"
DOWNLOAD_DIR = r"/path/to/your/folder"
ELEMENT_LOAD_WAIT_SEC = 30
DELAY_SEC = 3
MAIL_BOX = "Masterdata"
FULL_DATE_PATTERN =  r'(\d{4})-(\d{1,2})-(\d{1,2})'
STOP_DATE = "2023-12-16"

# Access the mailbox
def preprocess_element_class(x: str ):
    return x.replace(" ",".")

if __name__ == "__main__":

    with open('authen.json') as json_file:
        authen = json.load(json_file)

    chrome_options = Options()
    chrome_options.add_argument(f"--window-size={CHROME_WINDOW_SIZE}")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
    })


    driver = webdriver.Chrome(options=chrome_options)
    driver_wait = WebDriverWait(driver, ELEMENT_LOAD_WAIT_SEC)

    # Navigate to outlook.office.com/mail
    driver.get("https://outlook.office.com/mail/")

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

    time.sleep(DELAY_SEC)

    # Confirm my stay signed-in option
    stay_sign_in = (By.XPATH, "//button[@id='acceptButton']")
    stay_sign_in = driver_wait.until(EC.presence_of_element_located(stay_sign_in))
    stay_sign_in.click()

    time.sleep(DELAY_SEC)

    # Access the mailbox
    mail_box_panel = (By.CLASS_NAME, preprocess_element_class("C2IG3 if6B2 oTkSL iDEcr OPUpK"))
    mail_box_panel = driver_wait.until(EC.element_to_be_clickable(mail_box_panel))
    mail_box_panel.click()

    time.sleep(DELAY_SEC)

    mail_box = (By.CSS_SELECTOR, f"div[title^='{MAIL_BOX}']")
    mail_box = driver_wait.until(EC.presence_of_element_located(mail_box))
    mail_box.click()

    time.sleep(DELAY_SEC)

    # Select the targeted email.
    mail = (By.CLASS_NAME, preprocess_element_class("hcptT gDC9O"))
    mail = driver_wait.until(EC.presence_of_element_located(mail))

    while True:

        time.sleep(DELAY_SEC)
        
        # Check stop condition (email date)
        mail.click()
        email_date = (By.CLASS_NAME, preprocess_element_class("AL_OM sxdRi I1wdR"))
        email_date = driver_wait.until(EC.presence_of_element_located(email_date))
        email_date =  re.search(FULL_DATE_PATTERN,email_date.text,re.MULTILINE).group(0)
        if email_date == STOP_DATE:
            break

        time.sleep(DELAY_SEC)

        # Click on the attachment file
        attachment = (By.CLASS_NAME, "Y0d3P")
        attachment = driver_wait.until(EC.presence_of_element_located(attachment))
        attachment.click()

        time.sleep(DELAY_SEC)

        # Click on the download button
        download = (By.CSS_SELECTOR, 'button[name="ดาวน์โหลด"]')
        download = driver_wait.until(EC.presence_of_element_located(download))
        download.click()

        time.sleep(DELAY_SEC)

        # Return to the email list
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        download_popup = (By.CLASS_NAME, preprocess_element_class("m79Ne NjYhI"))
        driver_wait.until(EC.invisibility_of_element_located(download_popup))

        time.sleep(DELAY_SEC)

        # Move on to the next email
        mail.click()
        webdriver.ActionChains(driver).send_keys(Keys.DOWN).perform()
        mail = driver.switch_to.active_element