#TEST SolvSafety
import csv, re, time, openpyxl, os, time, datetime                                                                                #Importing Modules
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait

os.chdir(r'REMOVED')
filename = 'REMOVED.xlsx'
sheetname = 'Sheet1'
jobs = []

currentreport = openpyxl.load_workbook(filename)
sheet = currentreport[sheetname]
columnjoblookup = sheet['A']


#StartUp
driver = webdriver.Edge(r'REMOVED\msedgedriver.exe') #Start WebDriver to control Edge
driver.get('REMOVED')

for cell in columnjoblookup:  #loop to count number of jobs to look up and set upper bound on later loop                                                                                      
    jobs.append(cell.value)

print('REMOVED to Process = ' + str(len(jobs))) 

for i in range(2,len(jobs)+1):
#Refreshing Time Variables
    current_date = date.today() 
    now = datetime.datetime.now()
    current_time = now.strftime("%H:%M:%S")

    time.sleep(2)
    start = driver.find_element_by_xpath('REMOVED')
    start.click()
    driver.switch_to.window(driver.window_handles[0])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])




    #Incident Details
    REMOVED = driver.find_element_by_xpath('REMOVED')
    REMOVED.send_keys(sheet.cell(row=i, column=1).value)
    

    13/6/23 I have removed multiple further commands below this. There are 30-40 web form elements that are filled based on excel data.
    The script dealt with input text boxes, tick boxes, radio options and dropdown boxes. 
    There included a few dropdown boxes which had embedded "+" (expand) buttons that were more advanced to deal with, but I cannot include due to links to company webapps and HTML elements.
    

    REMOVED = driver.find_element_by_id(REMOVED')
    REMOVED.click()
    time.sleep(0.5)
    

    REMOVED_arrow = driver.find_element_by_id(REMOVED)
    REMOVED_arrow.click()
    time.sleep(0.5)
    REMOVED_Input = sheet.cell(row=i, column=7).value
    if REMOVED_Input == "REMOVED":
        REMOVED_Click = driver.find_element_by_xpath('REMOVED').click()
    elif REMOVED_Input == "REMOVED":
        driver.find_element_by_xpath('REMOVED').click()
    elif REMOVED_Input == "REMOVED":
        driver.find_element_by_xpath('REMOVED').click()
    else:
        print("REMOVED Error: Program has Failed")
    time.sleep(0.5)


    #ADDITIONAL DETAILS
    REMOVED


    #REPORTING DETAILS
    REMOVED


    #REMOVED DETAILS
    REMOVED


    #REMOVED INCIDENT SUPPORT
    REMOVED


    #ACTIONS DETAILS
    REMOVED
    First_Name_and_Surname_of_person_entering_Incident_Report = driver.find_element_by_xpath('REMOVED')
    First_Name_and_Surname_of_person_entering_Incident_Report.send_keys('Lachlan Young')
    time.sleep(1)

    #CLICK Save
    save = driver.find_element_by_id('REMOVED')
    time.sleep(1)
    save.click()
    time.sleep(30)
    REMOVED = driver.find_element_by_xpath('REMOVED').text
    print(REMOVED)
    sheet['REMOVED'+str(i)].value = REMOVED 
    time.sleep(1)
    Finish_Button = driver.find_element_by_xpath('REMOVED')
    Finish_Button.click()
    

