
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import selenium.webdriver.support.expected_conditions as EC
from selenium.webdriver.common.by import By
import time
from selenium.common.exceptions import TimeoutException
from openpyxl import load_workbook

# set file path
filepath=filename
# load excel-example.xlsx
wb=load_workbook(filepath)
# activate demo.xlsx
sheet=wb.active
# get b1 cell value

PATH = "/Users/ghost/Python/chromeDriver/chromedriver"
driver = webdriver.Chrome(PATH)

# Navigate to url

driver.get("https://www.redbubble.com/auth/login")

# Store google search box WebElement

driver.find_element_by_id("ReduxFormInput1").send_keys(sheet["A2"].value)
driver.find_element_by_id("ReduxFormInput2").send_keys(sheet["B2"].value)


delay = 120 # seconds
try:
    element = WebDriverWait(driver, 480).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='app']/div[1]/div/div[1]/div/div/div/header/div[3]/div[1]/div/div/div/div[1]/div/div/div/button"))
    )
    for i in range(3,10):
        rowC = 'C{0}'.format(i) #link
        rowD ='D{0}'.format(i)  #title
        rowE = 'E{0}'.format(i) #tags
        rowF = 'F{0}'.format(i) # image
        driver.get(sheet[rowC].value)
        driver.find_element(By.ID, "work_title_en").clear()
        driver.find_element(By.ID, "work_tag_field_en").clear()
        driver.find_element(By.ID, "work_title_en").send_keys(sheet[rowD].value)
        driver.find_element(By.ID, "work_tag_field_en").send_keys(sheet[rowE].value)
        driver.find_element(By.ID, "rightsDeclaration").click()
        driver.find_element(By.ID, "select-image-base").send_keys(sheet[rowF].value)
        time.sleep(30)
        driver.find_element(By.ID, "submit-work").click()
        time.sleep(30)
        driver.execute_script("window.open('about:blank', 'tab');")
        driver.switch_to.window("tab")
        url= sheet[rowC].value
        driver.get(url)
    

    
except:
    print("test")


#//*[@id='app']/div/div[1]/div[2]/div/div/div/div[1]/div/div

