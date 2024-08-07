from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import glob
import os
import shutil
from datetime import datetime

 
# Path to your WebDriver executable
driver_path = r"C:\Users\Administrator\Downloads\chromedriver-win64 (1)\chromedriver-win64\chromedriver.exe"
service=Service(executable_path=driver_path)
# Initialize WebDriver
driver = webdriver.Chrome(service=service)
driver.maximize_window()
# Navigate to a URL
driver.get('https://cma.opisnet.com/')
driver.find_element(By.XPATH,"//input[@id='edit-openid-connect-client-dow-jones-prod-idp-login']").click()
time.sleep(10)
#Providing credentials
driver.find_element(By.XPATH,"//input[@class='js-email-input']").send_keys("yaphanticha.j@indorama.net")
driver.find_element(By.XPATH,"//input[@class='password js-password-input']").send_keys("Indo062024!")
driver.find_element(By.XPATH,"//button[@class='solid-button basic-login-submit']//span[@class='text']").click()
 
time.sleep(10)

#Accept cookies
driver.find_element(By.XPATH,"//button[@class='agree-button eu-cookie-compliance-secondary-button button button--small btn btn-primary btn-xs  ']").click()
time.sleep(10)


driver.get('https://cma.opisnet.com/data-browser/buildQuery/Chemical%20Price%20&%20Economics')
time.sleep(20)
 
### If download notification is displayed should close
try:
    #Closing download notification
    close=driver.find_element(By.XPATH,"//button[@class='toast-notification__close-button']")
    close.click()
    print('Download notification closed')
except Exception:
    pass
 
 
#Navigating to 'My Saved' checkbox
driver.find_element(By.XPATH,"//a[@href='/data-browser/mySaved']").click()
time.sleep(15)
print("Navigated to my-saved option")

#Navigating to CMA_Monthly dataset checkbox
search_bar=driver.find_element(By.XPATH,"//a[contains(@href,'/data-browser/buildQuery/queryId/')]")
search_bar.click()
time.sleep(15)
 
## Clicking on export checkbox
export_button=driver.find_element(By.XPATH,"//button[text()='Export']")
count=0    
while count==0:
    try:
        export_button.click()
        count=1
    except Exception:
        count=0
print("Export is enabled")
 
time.sleep(10)
 
## Proccesing to click Excel Static option in Export
nxt_button=driver.find_element(By.XPATH,"//button[text()='Excel Static']") 
nxt_button.click()
 
print("Selected option for saving as excel file")
 
## Wait for request for file to be downloaded
time.sleep(120)
 
### Clicking on notif link to download
down=driver.find_element(By.XPATH,"//a[@class='toast-notification__url']")
print("Downloading notification popped")
 
count=0    
while count==0:
    try:
        down.click()
        count=1
    except Exception:
        count=0
 
print("File is downloaded")
 
time.sleep(30)
 
print(driver.title)  # Prints the title of the page
 
# Close the browser
driver.quit()

# #move file to particular sharepoint folder
# path_to_files = r'C:\Users\saurabh.g\Downloads' 
  
# pattern = os.path.join(path_to_files, 'Chemical Price & Economics__*.xlsx')

# matching_files = [i for i in glob.glob(pattern)]

# file=os.path.basename(matching_files[0])

# file=file.rsplit('.xlsx')[0]+' monthly'

# # file=file.rsplit('__',maxsplit=1)[0]
# new_file = rf'C:\Users\saurabh.g\OneDrive - Indorama Ventures PCL\Project LightHouse\Archival\CMA\{file}.xlsx'
# shutil.move(matching_files[0], new_file)


## # Moving file from CMA to Archived

cma_folder = r'C:\Users\Administrator\OneDrive - Indorama Ventures PCL\Project LightHouse\Archival\CMA'
destination_folder_for_move = r'C:\Users\Administrator\OneDrive - Indorama Ventures PCL\Project LightHouse\Archival\CMA\Archived'
pattern_in_CMA = os.path.join(cma_folder, 'Chemical Price & Economics*monthly.xlsx')
matching_files_in_CMA = glob.glob(pattern_in_CMA)

for file_path in matching_files_in_CMA:
    moved_file_destination = os.path.join(destination_folder_for_move, os.path.basename(file_path))


    if os.path.exists(moved_file_destination):
        os.remove(moved_file_destination)

    try:
        shutil.move(file_path,destination_folder_for_move)
        print(f'moved file {file_path} to {destination_folder_for_move}')
    except Exception as e:
        print((f'failed to move {file_path} to {destination_folder_for_move}'))

# Moving file from downlods to CMA folder

path_to_files = r'C:\Users\Administrator\Downloads' 
pattern = os.path.join(path_to_files, 'Chemical Price & Economics__*.xlsx')
matching_files = glob.glob(pattern)

file = os.path.basename(matching_files[0])
    
file_base = file.rsplit('.xlsx', 1)[0]

new_file_name = f"{file_base} monthly.xlsx"

new_folder = rf'C:\Users\Administrator\OneDrive - Indorama Ventures PCL\Project LightHouse\Archival\CMA\\'

try:
        if os.path.exists(os.path.join(new_folder, new_file_name)):
            os.remove(os.path.join(new_folder, new_file_name))
            print(f"File deleted from CMA: {new_file_name}")
        
        source_file = os.path.join(path_to_files, file)
        renamed_source_file = os.path.join(path_to_files, new_file_name)
        os.rename(source_file, renamed_source_file)
        
        shutil.move(renamed_source_file, new_folder)
        print(f'Moved file {new_file_name} to {new_folder}')
except Exception as e:
    print(f'Failed to move {new_file_name} to {new_folder}: {e}')