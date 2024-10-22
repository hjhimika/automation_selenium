from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Setup Chrome with a specific user profile
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=/path/to/your/custom/profile")
options.add_argument("--disable-popup-blocking")
# Initialize WebDriver with service to manage versions
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
driver.get("https://outlook.office.com/mail/")
# Additional steps can be added here based on specifics of the pop-up
# Handling more elements, logging in, etc.