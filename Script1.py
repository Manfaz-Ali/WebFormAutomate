import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Open the Excel file and get the worksheet
workbook = openpyxl.load_workbook('data.xlsx')
worksheet = workbook['Sheet1']

# Get the login credentials from the worksheet
url = worksheet['B49'].value
user_nm = worksheet['B50'].value
password = worksheet['B51'].value

# Split the combined value into username and password
user_id = user_nm

# Set the Accept-Language header to English
options = webdriver.ChromeOptions()
options.add_argument('lang=en')

# Start the Chrome webdriver and navigate to the login page
driver = webdriver.Chrome(executable_path="chromedriver.exe",options=options)
driver.get(url)

# Wait for the login form to appear
wait = WebDriverWait(driver, 10)
login_form = wait.until(EC.presence_of_element_located((By.ID, 'acceso')))

print(user_id)
print(user_nm)
print(password)

# Find the form fields by name within the login form element
security_code_field = driver.find_element(By.NAME,'colaborador')
security_code_field.send_keys(user_id)
username_field = driver.find_element(By.NAME,'alias')
username_field.send_keys(user_nm)
password_field = driver.find_element(By.NAME,'password')
password_field.send_keys(password)

time.sleep(10)
submit_btn = driver.find_element('id','accesoLogin')
submit_btn.click()



# Wait for the dashboard page to load
dashboard = wait.until(EC.title_contains('Dashboard'))

# Do something on the dashboard page, e.g. scrape some data or click some buttons
print("You login successfully")
# Close the browser window
driver.quit()
