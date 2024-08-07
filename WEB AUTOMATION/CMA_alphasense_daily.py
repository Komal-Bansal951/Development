from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui as pg
 
chrome_options = Options()
 
### Suppress printer to switch to Save As option
chrome_options.add_argument('--kiosk-printing')
 
# Path to your WebDriver executable
driver_path = r"C:\Users\Administrator\Downloads\chromedriver-win64 (1)\chromedriver-win64\chromedriver.exe"
service=Service(executable_path=driver_path)
 
# Initialize WebDriver
driver = webdriver.Chrome(options=chrome_options)
driver.maximize_window()
 
# Navigate to a URL
driver.get('https://cma.opisnet.com/')
driver.find_element(By.XPATH,"//input[@id='edit-openid-connect-client-dow-jones-prod-idp-login']").click()
time.sleep(5)
 
#Providing credentials
driver.find_element(By.XPATH,"//input[@class='js-email-input']").send_keys("yaphanticha.j@indorama.net")
driver.find_element(By.XPATH,"//input[@class='password js-password-input']").send_keys("Indo062024!")
driver.find_element(By.XPATH,"//button[@class='solid-button basic-login-submit']//span[@class='text']").click()
time.sleep(10)
 
#Accept cookies
driver.find_element(By.XPATH,"//button[@class='agree-button eu-cookie-compliance-secondary-button button button--small btn btn-primary btn-xs  ']").click()
time.sleep(10)


##NAM Aromatics
 
driver.get('https://cma.opisnet.com/publications')
time.sleep(5)
 
### If download notification is displayed should close
# try:
#     #Closing download notification
#     close=driver.find_element(By.XPATH,"//button[@class='toast-notification__close-button']")
#     close.click()
#     print('Download notification closed')
# except Exception:
#     pass
 
# Locate and interact with the search box
search_box = driver.find_element(By.XPATH, "//input[@id='edit-search-api-fulltext--3']")
#print("Search box is available")
time.sleep(2)
search_text = "North America Aromatics daily"
search_box.send_keys(search_text)
time.sleep(3)
 
driver.find_element(By.XPATH, "//input[@id='edit-submit-acquia-search']").click()
time.sleep(3)
 
file_button = driver.find_element(By.XPATH,"//a[contains(@href,'/report/daily-north-america-aromatics-report-issue')]")
file_name = file_button.text

count=0    
while count==0:
    try:
        file_button.click()
        count=1
    except Exception:
        count=0
 
time.sleep(3)
 
## Interact with the print option to save file
print_button=driver.find_element(By.XPATH,"//div[@class='page-actions']//button[@class='icon-action page-actions__action-button']")
print_button.click()
 
### Time to create visibilty of print dialog box
time.sleep(5)
#Provide the file name to be saved
pg.typewrite(rf"C:\Users\Administrator\OneDrive - Indorama Ventures PCL\01 Industry Report\CMA\NAM Aromatics (Daily)\2024\{file_name}")
pg.press('enter')
 
time.sleep(15)
print("NAM Aromatics daily file exported")