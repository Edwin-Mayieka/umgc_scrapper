from selenium import webdriver
import time
import re
import datetime
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import openpyxl
import os
# Set the path to the Chrome binary
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = '/usr/bin/chromium'
#chrome_options.binary_location = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 5)

home_page='https://www.umgc.edu/login'
login_email='cnanakumo@student.umgc.edu'
login_pswd='Samhilda@2023'
class_code=708985 #adjust accordingly
topic_code=3860892 #adjust accordingly

discussion_url_unread=f'https://learn.umgc.edu/d2l/le/{class_code}/discussions/topics/{topic_code}/View?filters=unread&groupFilterOption=0'
discussion_url_all=f'https://learn.umgc.edu/d2l/le/{class_code}/discussions/topics/{topic_code}/View?groupFilterOption=0'

def startPostNav(startPost):
    wait = WebDriverWait(driver, 7)
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.d2l-iterator-button-next')))
        
        for post in range(startPost):
            next_post=driver.find_element(By.CLASS_NAME,'d2l-iterator-button-next').get_attribute('href')
            driver.get(next_post)
            time.sleep(5)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.d2l-user-profile-handle-normal')))

    except:
        print("Error navigating to start page*")


              
def nextPage():
    wait = WebDriverWait(driver, 7)
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.d2l-iterator-button-next')))
        next_post=driver.find_element(By.CLASS_NAME,'d2l-iterator-button-next').get_attribute('href')
        if next_post:
                driver.get(next_post)
                time.sleep(5)
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.d2l-user-profile-handle-normal')))

    except:
        print("Error navigating to next pages*")

def print_data_to_excel(name, email):
    current_dir = os.getcwd()
    file_path = os.path.join(current_dir, "ugcm.xlsx")

    # Create the file if it doesn't exist
    if not os.path.exists(file_path):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.cell(row=1, column=1).value = "Name"
        sheet.cell(row=1, column=2).value = "Email"
        workbook.save(file_path)

    # Open the existing workbook
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Get the last row
    last_row = sheet.max_row + 1

    # Print data to respective columns starting from the last row
    sheet.cell(row=last_row, column=1).value = name
    sheet.cell(row=last_row, column=2).value = email

    # Save the workbook
    workbook.save(file_path)
    print(f"Data printed to Excel successfully at: {file_path}")

def navigate():
    startPost=None
    print("Script starting***")
    posts_filter=input("Enter Posts Filter: \n 1.ALL  2.Unread: ")
    posts_filter=int(posts_filter)
    if posts_filter ==1 :
        startPost=input("Start at Post Number:  ")
        startPost=int(startPost)
    
    posts2read=input("Enter Number of posts to read: ")
    posts2read=int(posts2read)



    driver.get(home_page)
    time.sleep(3)
    
   
    
    # login
    login_page=driver.find_element(By.ID, 'button-a198d4e78c').get_attribute('href')
    driver.get(login_page)
    time.sleep(3)
    WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.ID, 'idSIButton9')))
    email_field = driver.find_element(By.ID, 'i0116')
    email_field.send_keys(login_email)
    to_passwd_btn=driver.find_element(By.ID, 'idSIButton9')
    time.sleep(3)
    to_passwd_btn.click()
    time.sleep(3)
    password_field = driver.find_element(By.ID, 'i0118')
    password_field.send_keys(login_pswd)
    time.sleep(3)
    submit_btn = driver.find_element(By.CSS_SELECTOR, 'input#idSIButton9')
    submit_btn.click()
    #probably add code to check if certain tag exists before navigating
    print("login done")
    time.sleep(3)
    if posts_filter==1:
        driver.get(discussion_url_all)
    else:
         driver.get(discussion_url_unread)

    
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a.d2l-linkheading-link')))
    post_heading=driver.find_element(By.CLASS_NAME, 'd2l-linkheading-link') #gets first post heading link
    
    time.sleep(5)
    post_heading.click() #clicks on first link
    # test_url='https://learn.umgc.edu/d2l/le/708985/discussions/threads/28379277/View?searchText=testing+the+waters'
    # driver.get(test_url)
    time.sleep(3)
    if startPost != 1 :
        startPostNav(startPost)
    if posts2read <= 1 :
        user_ids=get_user_ids()
        getContacts(user_ids)
    else:
        for post2read in range(posts2read):
            user_ids=get_user_ids()
            getContacts(user_ids)
            nextPage()

def get_user_ids():
    ids=set()
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.d2l-user-profile-handle-normal')))
    user_avartars = driver.find_elements(By.CLASS_NAME,'d2l-user-profile-handle-normal')
    print(f"no of avatars: {len(user_avartars)}")
    
    
    for avatar in reversed(user_avartars):
        # print(f"avartar id: {avatar.get_attribute('id')}")
        avatar.click() #maybe wait until certain elements of the avarter are visible before proceeding to ensure that the placeholder in the DOM
        time.sleep(7)
     
    id_elements=driver.find_elements(By.CSS_SELECTOR, '.d2l-placeholder.d2l-placeholder-inner.d2l-placeholder-live')
    for id_element in id_elements:
        placeholder_id=id_element.get_attribute('id')
        if 'profilePlaceholder' in placeholder_id:
            ids.add(placeholder_id.replace('profilePlaceholder',''))
    print(f"ids: {ids}")
    return ids
   

def getContacts(ids): 
   
    for id in ids:
        driver.execute_script(f"window.open('https://learn.umgc.edu/d2l/lms/email/integration/AdaptLegacyPopupData.d2l?ou={class_code}&p=0&ext=1&cb=d2l_2_0_637&singleUserId={id}', 'new_window');")
        window_handles = driver.window_handles
        # Switch to the second window
        new_window = window_handles[1]
        driver.switch_to.window(new_window)
        time.sleep(10)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.d2l-multiselect-choice')))
        element=driver.find_element(By.CSS_SELECTOR, '.d2l-multiselect-choice') #d2l-heading vui-heading-1
        contact=element.text
        print(contact)
        match = re.match(r'"(.*?)" <(.*?)>', contact)
        if match:
            name=match.group(1)
            email=match.group(2).replace('<', '').replace('>', '')
            print_data_to_excel(name,email)
        else:
            return print("regex not working")
        driver.close()

    # Switch to the first window
        first_window = window_handles[0]
        driver.switch_to.window(first_window)
    



navigate()




