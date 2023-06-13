# PROGRAM: Copy Job Note into Excel

## KNOWN ISSUES
## 1. If REMOVED job, program fails - non-issue, boss said we don't need to include
## 2. Incomplete note if there is a colon in job note


##Possible Improvements
## 1. Resolve Known Issue 2, perhaps with a split across next lines or id of REMOVED (eg: REMOVED)
## 2. Set an Input Validation up for Enter your Inputs, especially the spreadsheet data
## 3. Break into Chunks. REMOVED slows down once it gets past a certain number. After every 100 jobs, restart REMOVED screen.

#SetUp
import re, time, openpyxl, os                                                                                #Importing Modules
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
jobs = []                                                                                                               #Set up empty list for second for loop

os.chdir(r'REMOVED')                     #Set Directory

#Enter your Input Variables
filename = 'REMOVED.xlsx'           #must include xlsx and update lasy line
sheetname = 'REMOVED'
columnletter = 'REMOVED'  #This is for job number column
LastJobNote = 'REMOVED'
DateOfLastJobNote = 'REMOVED'
LastOnsiteDate = 'REMOVED'
REMOVED = 'REMOVED'
REMOVED = 'REMOVED'

#Start Up Reminder to Validate Inputs
time.sleep(1)
print('Please ensure your inputs are correct by doing the following \n')
time.sleep(1)
print('1. Directory is correct \n')
time.sleep(1)
print('2. You have the correct filename and it is copied into the quotes on the last line \n')
time.sleep(1)
print('3. You have the correct sheetname and column that the job numbers are in \n')
time.sleep(1)
print('4. You have the correct columns for output data \n')
time.sleep(1)
print('5. Make sure your for loop has the correct range \n')
time.sleep(1)
print('Please close the excel file you are processing')
time.sleep(5)

                                                       
#Output Variables
currentreport = openpyxl.load_workbook(filename)
sheet = currentreport[sheetname]
columnjoblookup = sheet[columnletter]

for cell in columnjoblookup:  #loop to count number of jobs to look up and set upper bound on later loop                                                                                      
    jobs.append(cell.value)

#Code for user to see results
print('Jobs to Process = ' + str(len(jobs)))                                                                            
print('Jobs processed...')

def webdriver_start():
    global driver
    driver = webdriver.Edge(r'REMOVED\msedgedriver.exe') #Start WebDriver to control Edge
    driver.get('REMOVED')                          #Open REMOVED Screen                                                                                                      



#Create the Excel Loop to look up all job numbers
webdriver_start()
driver.implicitly_wait(30)
for i in range(2,len(jobs)+1):
    try:
        if i % 100 == 0:
            driver.quit()
            time.sleep(2)
            webdriver_start()
            pass

        REMOVED = ''                                                                                                #Reset REMOVED Content and REMOVED to avoid duplicates
        REMOVED = ''
        REMOVED=' '                                                                                                 
        jobNumber = sheet.cell(row=i, column=1).value                                                                        #assigning job number variable to current job number value going through loop
                                                                                                                             #print current job procesed
        #Job Number & Job Notes LookUp
        REMOVED = driver.find_element_by_xpath('REMOVED')                                                		#identify REMOVED in REMOVED Screen's HTML and put it under variable job search
        REMOVED.send_keys(jobNumber, Keys.ENTER)                                                                         #enter job number and REMOVED
        time.sleep(1)
        JobNotes = driver.find_element_by_xpath('REMOVED').text                                             		#identify job notes in REMOVED HTML and put it under variable Job Notes
        time.sleep(0.5)
        # Extract Last Update & Clean It
        REMOVEDdate = re.split('([0-9]{2}/[0-9]{2}/[0-9]{4})', JobNotes)                                                      #split all job notes at the time spot                                               
        REMOVED_date = splitdate[-2:]                                                                                        #create a slice of the last REMOVED and date of REMOVED (isolate last REMOVED and date)
        Cleannote2 = re.split(':', REMOVED_date[1])                                                                          #split on colon to make date and REMOVED different items

        # Extracting Last Onsite and Upgrade & Clean It
        events = driver.find_element_by_xpath('REMOVED').text
        time.sleep(0.5)
        eventslist = re.split('\n',events)
        REMOVED = [x for x in eventslist if re.search('REMOVED',x)]
        REMOVED = [y for y in eventslist if re.search('REMOVED',y)]
        for x in range(len(REMOVED)):
            REMOVED = re.split(' ', REMOVED[x])
            REMOVEDcontent = " ".join(REMOVED[0:2])
            REMOVEDdates = " ".join(REMOVED[2:5])
        for y in range(len(REMOVED)):
            splitREMOVED = re.split(' ', OnSite[y])
            REMOVEDdates = " ".join(splitOnSite[2:5])

        # Paste Update and Date into Correct Excel Cells
        sheet[str(REMOVED)+str(i)].value = REMOVED[0]
        sheet[str(REMOVED)+str(i)].value = Cleannote2[-1]
        sheet[str(REMOVED)+str(i)].value = REMOVEDdates
        sheet[str(REMOVED)+str(i)].value = REMOVEDcontent
        sheet[str(REMOVED)+str(i)].value = REMOVEDdates
        currentreport.save("25.10 REMOVED.xlsx")                                                                                 #Save report
        print(str(i-1) + ': ' + str(jobNumber))

    except:
        print("ERROR: Notes not added. " + str(i-1) + ': ' + str(jobNumber))
        sheet[str(DateOfLastREMOVED)+str(i)].value = 'ERROR: Input Failed'


    







