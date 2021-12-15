from selenium import webdriver
import time
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2 import service_account
from tqdm import trange
import os
from progress.bar import Bar
from datetime import datetime

'''
Don't forget, this script requires the key.json file to be in its working directory.
'''

os.getcwd()
dir = os.chdir(dir)

username = input('What is your username?')
password = input('What is your password?')
keysJSON = 'path to JSON file'

# Delete the file
try:
    os.remove('results-survey683435.xlsx')
except: 
    pass

# Call the web browser
chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : dir} # Changes the download directory
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument("--window-size=1920, 1080")
chrome_options.add_argument("--headless")
browser = webdriver.Chrome('chromedriver', options=chrome_options)

# Navigate to the chosen website 
try:
    browser.get('https://survey.sae.edu/index.php/admin/export/sa/exportresults/surveyid/683435')
    time.sleep(2)
except:
    pass
   
# This downloads the data
bar = Bar('Downloading file...', max=30)
def exportdata():
    try:
        browser.find_element_by_css_selector('#user').send_keys(username)
        browser.find_element_by_css_selector('#password').send_keys(password)
        browser.find_element_by_xpath('//*[@id="loginform"]/div[2]/div/p/button').click()
        browser.find_element_by_css_selector('#xls').click()
        browser.find_element_by_css_selector('#panel-4 > div.panel-body > div:nth-child(1) > div > label:nth-child(4)').click()
        submit = browser.find_element_by_css_selector('#export-button')
        browser.minimize_window()
        submit.click()
    except:
        pass
        
# Function executes
exportdata()

# Bar visual
for i in range(30):
    time.sleep(1)
    bar.next()

browser.quit()

# Read & filter the downloaded file
print("Filtering...")
time.sleep(2) # Wait 2 seconds
try:
    df = pd.read_excel('results-survey683435.xlsx')
    df = df[df['TP. Teaching Period'].isin(["20T3","21T1","21T2"]) & df['Campus. Campus'].isin(["Liverpool", "London", "Oxford", "Glasgow","Online"])].dropna(how='all')
except:
    pass

# Google Sheets Setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = keysJSON
credentials = None
credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID and range of a sample spreadsheet.
modulesurveydata = '1UCYodN9q1MYt3embI-oelo9994ywSpYolnkGduzA4JM' ## Change this to the new Survey Tracker
service = build('sheets', 'v4', credentials=credentials)

# Call the Sheets API and write data
df.fillna('', inplace=True)
sheet = service.spreadsheets()
data = df.values.tolist()
def updatedata():
    print("Uploading data...")
    sheet.values().update(spreadsheetId=modulesurveydata, range="survey data!A2", valueInputOption="USER_ENTERED", body={"values":data}).execute()
    print("Done.")

# Function executes
updatedata()

# Delete the file
os.remove('results-survey683435.xlsx')