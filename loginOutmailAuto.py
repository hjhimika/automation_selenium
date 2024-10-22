from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import time

def login_outlook(email, password):
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get("https://outlook.live.com/owa/")
    
    # Click on 'Sign in' button
    sign_in_button = driver.find_element(By.LINK_TEXT, "Sign in")
    sign_in_button.click()
    time.sleep(2)

    # Enter email and proceed
    email_input = driver.find_element(By.NAME, "loginfmt")
    email_input.send_keys(email)
    email_input.send_keys(Keys.RETURN)
    time.sleep(2)

    # Enter password
    password_input = driver.find_element(By.NAME, "passwd")
    password_input.send_keys(password)
    password_input.send_keys(Keys.RETURN)
    time.sleep(2)

    # Confirm login
    driver.find_element(By.ID, "idSIButton9").click()
    time.sleep(3)

    return driver
