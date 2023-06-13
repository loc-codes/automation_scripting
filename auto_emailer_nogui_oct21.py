#Module Set Up
import pandas as pd, openpyxl, win32com.client as client, time, datetime,os

#Excel Variables
print('Check that the last line is mail.Save() before you proceed')
time.sleep(2)
path = input('Please enter the path: ')
file = input('Please enter .xlsx file: ')
brand = input('Please enter the brand you would like in the email subject: ')
   
#Email Variables
start = '''<div class=WordSection1>
<p class=MsoNormal>Hi Team,<br>
<br>
REMOVED
</div>'''

signature ='''<html><head>
...REMOVED...
 <o:DocumentProperties>
  <o:Author>Lachlan Young</o:Author>
  <o:Template>Normal</o:Template>
  <o:LastAuthor>Lachlan Young</o:LastAuthor>
  <o:Revision>1</o:Revision>
  <o:TotalTime>1</o:TotalTime>
  <o:Created>2021-10-11T04:29:00Z</o:Created>
  <o:LastSaved>2021-10-11T04:30:00Z</o:LastSaved>
  <o:Pages>1</o:Pages>
  <o:Words>80</o:Words>
  <o:Characters>457</o:Characters>
  <o:Lines>3</o:Lines>
  <o:Paragraphs>1</o:Paragraphs>
  <o:CharactersWithSpaces>536</o:CharactersWithSpaces>
  <o:Version>16.00</o:Version>
 </o:DocumentProperties>
<style>
REMOVED
</style>
<style>
REMOVED
</style>
REMOVED
</head>

<body ...REMOVED>
REMOVED
</body>
REMOVED
</body>
REMOVED
</html>
'''


##Pandas Set Up and Grabbing emailAddress and contractor variables
df = pd.read_excel(path+'\\'+file)                                                                                                           #Pandas Reads Excel
resources_df = df[['REMOVED','REMOVED']]                                                                                         #Creates table of Resources and their emails
resources_df = resources_df.drop_duplicates()                                                                                           #Removes duplicate
resources_df = resources_df.reset_index(drop=True)                                                                                      #Removes numerical index
emailAddress = list(resources_df['Email'])                                                                                              #Variable 1: Email Address
contractor = list(resources_df['REMOVED'])                                                                                     #Variable 2: Contractor
grouped_resource = df.groupby(['REMOVED'])[['REMOVED',...,'REMOVED']]  #Creates individual tables by Contractor

##Generating Individual Tables per Contractor
emailtable = []                                                                                                                         #Variable 3: List of Tables for emails, currently empty
           
for resource in contractor:                                                                                                             #For Loop
    input_df = grouped_resource.get_group(resource)                                                                                     #Creates a new table based on Contractor
    input_df = input_df.style.set_table_styles([{'selector': 'th','props': [('background-color', 'lightpink'),('font-family', 'Calibri, sans-serif'),('border-color', 'black'),('border', 'solid'),('border-width', '1.3px')]}]).set_properties(**{'color': 'black', 'background':'lightpink', "border-collapse": "collapse" ,"border": "1.3px solid black",'font-family': 'Calibri, sans-serif'})
    input_df.hide_index().render()                                                                                                      #Hides numerical index
    if input_df not in emailtable:                                                                                                      #Converts to html and adds to list if not already in there
        html = input_df.to_html()
        emailtable.append(html)

##Email Function
def send_email(x):                                                                                                                      #Define email function                                  
    REMOVED = client.Dispatch('REMOVED')                                                                                    #Set up outlook by assigning variable
    mail = REMOVED.CreateItem(0)                                                                                                        #set up message variable as blank email 
    mail.To = emailAddress[x]                                                                                                           #looped email address, matched with correct contractor name and table based on number i     
    mail.SentOnBehalfOfName = 'REMOVED@city-holdings.com.au'                                                                          #Sets email to send from closedown email  
    mail.Subject = 'Open ' + str(brand) + ' Jobs ' + str(datetime.date.today().strftime('%d/%m')) + ' - ' + contractor[x]
    mail.HTMLBody =  start + emailtable[x] + "<br>" + "<br>" + signature
    mail.Send()                                                                                                                         #saves email to drafts

#For Loop for Individualised Emails
for x in range(len(contractor)):
    send_email(x)




#Possible Improvements
#1: instead of generic: "Hi Team", have persons name



