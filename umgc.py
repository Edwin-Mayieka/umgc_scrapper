from selenium import webdriver
import time
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
home_page='https://www.umgc.edu/login'
class_code=708985 #adjust accordingly
topic_code=3860892 #adjust accordingly
discussion_url=f'https://learn.umgc.edu/d2l/le/{class_code}/discussions/topics/{topic_code}/View?filters=unread&groupFilterOption=0'
course_url=f'https://app.schoology.com/course/{class_code}/members'
# excel file name
current_time = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')


# Set up the Chrome driver with the downloaded ChromeDriver executable

#= Navigate to the Schoology login page


emails = []
phones=[]
names=[]

def loadData(page):
    global currentPage
    # print(f"loadData round: {page+1}")
    wait = WebDriverWait(driver, 10)
    time.sleep(5)
    table = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'table[role="presentation"]')))
    ids=[]
    rows = table.find_elements(By.TAG_NAME, 'tr')
    for row in rows:
        user_id = row.get_attribute('id')
        ids.append(user_id)
        element = row.find_element(By.CSS_SELECTOR, 'td.user-name a.sExtlink-processed')
        name = element.text
        #last_name = element.find_element(By.TAG_NAME, 'b').text
        names.append(name)
        #print(name)

    for id in ids:
        if id:
           
            # Construct the user page URL using the extracted user ID
            user_page_url = f'https://app.schoology.com/user/{id}/info' 
            driver.get(user_page_url)
            # Find the email element and extract the email address
            try:
                email_element = driver.find_element(By.CSS_SELECTOR, 'a.sExtlink-processed.mailto')
                email = email_element.get_attribute('href').replace('mailto:', '')
                emails.append(email)
            except NoSuchElementException:
                    emails.append('')
                   
            try:
                phone_element = driver.find_element(By.XPATH, "//td/a[@class='sExtlink-processed']")
                if phone_element is not None:
                    phone = phone_element.get_attribute('href').replace('tel:', '')
                    phones.append(phone)
  
                else:
                    phone = ''
                    phones.append(phone) 
                
            except NoSuchElementException:
                    phones.append('')
    if page!=pages-1:                      
        driver.get(course_url)
        time.sleep(5)
        try:
            # print("Going to currentPage")
            wait = WebDriverWait(driver, 5)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.next')))
            startPageNav(currentPage)
            nextPage()
            print(f" Now on page>> {currentPage}")
            currentPage += 1
            
        except:
            print("Error resuming current page")



def startPageNav(startPage):
    global currentPage
    wait = WebDriverWait(driver, 10)
    try:
        print("Going to startPage")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.next')))
        for page in range(startPage-1): #changed startPage-1
            next_page = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.next')))
            if next_page:
                next_page.click()
                time.sleep(5)
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.next')))
        print(f" Now on page> {currentPage}")
    except (ValueError,TypeError) as e:
        print("Error going to start page",e)


              
def nextPage():
    wait = WebDriverWait(driver, 7)
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.next')))
        next_page = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.next')))
        if next_page:
                next_page.click()
                print("Clicked next.")
                time.sleep(5)
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.next')))
    except:
        print("No more next pages*")

def print_data_to_excel(names, emails, phones, filename):
    current_dir = os.getcwd()
    file_path = os.path.join(current_dir, filename)
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Add headers
    sheet.cell(row=1, column=1).value = "Names"
    sheet.cell(row=1, column=2).value = "Emails"
    sheet.cell(row=1, column=3).value = "Phone"

    # Print data to respective columns
    for i in range(len(names)):
        sheet.cell(row=i+2, column=1).value = names[i]
        sheet.cell(row=i+2, column=2).value = emails[i]
        sheet.cell(row=i+2, column=3).value =phones[i]


    # Save the workbook
    workbook.save(file_path)
    print(f"Data printed to Excel successfully at:{current_time}")


def getData():
    print("Script starting***")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 5)
    driver.get(home_page)
    time.sleep(3)
    
    
    # pageToStart=input("Enter startPage: ")
    # pages=input("Enter the number of pages to go through. ie pages=totalScholars/ScholarsPerPage : ")
    # pageToStart=int(pageToStart)
    # currentPage=pageToStart
    # pages=int(pages)
    #filename='courseCode_'+str(class_code)+'_startPage_'+str(pageToStart)+'_'+ current_time + '.xlsx'
    
    # login
    login_page=driver.find_element(By.ID, 'button-a198d4e78c').get_attribute('href')
    driver.get(login_page)
    time.sleep(3)
    email_field = driver.find_element(By.ID, 'i0116')
    email_field.send_keys('cnanakumo@student.umgc.edu')
    time.sleep(3)
    to_passwd_btn=driver.find_element(By.ID, 'idSIButton9')
    to_passwd_btn.click()
    password_field = driver.find_element(By.ID, 'i0118')
    password_field.send_keys('Samhilda@2023')
    wait.until(EC.presence_of_element_located((By.ID, 'idSIButton9')))
    submit_btn = driver.find_element(By.ID, 'idSIButton9')
    wait.until(EC.presence_of_element_located((By.ID, 'idSIButton9')))
    submit_btn.click()
    #probably add code to check if certain tag exists before navigating
    print("login done")
    time.sleep(5)
    driver.get(discussion_url)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a.d2l-linkheading-link')))
    post_heading=driver.find_element(By.CLASS_NAME, 'd2l-linkheading-link')
    # postHeading=post_headings[0].get_attribute('href')
    # print(post_heading.get_attribute('href'))
    time.sleep(5)
    post_heading.click()
    time.sleep(3)
    user_id_string=driver.find_element_by_class_name('d2l-placeholder d2l-placeholder-inner d2l-placeholder-live').get_attribute('id')
    print(f"user id: {user_id_string}")
    
    driver.execute_script("window.open('https://learn.umgc.edu/d2l/lms/email/integration/AdaptLegacyPopupData.d2l?ou=708985&p=0&ext=1&cb=d2l_2_0_637&singleUserId=476920', 'new_window');")
    window_handles = driver.window_handles
    # Switch to the second window
    new_window = window_handles[1]
    driver.switch_to.window(new_window)
    time.sleep(5)
    driver.close()
# Switch to the first window
    first_window = window_handles[0]
    driver.switch_to.window(first_window)
    # driver.close()
   

    
    
    # time.sleep(3) 
 
    # Submit the login form
    # driver.find_element(By.ID, 'edit-submit').click()
    # driver.implicitly_wait(10)
    # driver.get(course_url)
    # if pageToStart==1:
    #     for page in range(pages):
    #         loadData(page)
    # else:
    #     startPageNav(page)
    #     time.sleep(5)
    #     for page in range(pages):
    #         loadData(page)      
    # print_data_to_excel(names,emails,phones,filename=filename)
    # Quit
    # driver.quit()


getData()




