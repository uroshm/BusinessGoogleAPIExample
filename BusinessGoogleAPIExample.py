
from __future__ import print_function
from os import system, name

import json
import gspread
import sys
import os
import time
import email
import string
import base64
import imaplib
import ConfigParser
import datetime
import pickle
import os.path

from datetime import date, timedelta
from email.parser import FeedParser

from pprint import pprint

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient import discovery

from oauth2client.service_account import ServiceAccountCredentials

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

def createGoogleSheet(title):
    # use creds to create a client to interact with the Google Drive API
    # We will be using gspread lib here to interact with Google Sheets.
    retVal = []
    
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)

    sheet = client.create(title)
    wks = sheet.worksheet("Sheet1")
    wks.update_acell("a1","Customer Complaint:")
    wks.update_acell("a5","Customer Name:")
    wks.update_acell("a9","Cusotomer Address:")
    wks.update_acell("c22","=c17-c18-c19+c20-c21")
    
    #This is very important - with whom do we want to share this google sheet with? Business rules apply here:
    sheet.share("BusinessXYZ@gmail.com", perm_type="user", role="writer")
    retVal.append('https://docs.google.com/spreadsheets/d/'+sheet.id+'/edit#gid=0')
    retVal.append(sheet.id)
    return retVal
    

def readEmail(username, pasProvider2rd, sender_of_interest):
    print("BusinessXYZScan: beginning to scan for sender: "+sender_of_interest)
    
    conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    conn.login(username, pasProvider2rd)
    conn.select('INBOX')
    conn.readonly = True
    countEmails = 0
    
    # Let's read our inbox for the time period (today minus three days)
    todayDt = date.today()-timedelta(days=3)    
    todayStr = todayDt.strftime('%d-%b-%Y')
        
    seconds = time.time()
    secondsSearch = str(int(seconds - (3600*4)))
    print(secondsSearch) 
    bodyList = []
    
    #using conn.uid, we will search our gmail based on our sender of interest as well as the date range.
    result, data = conn.uid('search','(HEADER FROM "' + sender_of_interest + '")','(SINCE '+todayStr+')')
    
    for num in data[0].split():
        countEmails = countEmails + 1
        result, email_data = conn.uid('fetch', num, '(RFC822)')
        raw_email = email_data[0][1]
        
        #It is possible that we will be getting emails with different types of text encryption. We may need to hanlde this on a one-off basis.
        #Below, I will use base64 decryption in the first case and utf-8 in the second.

        if(sender_of_interest=="confirm@xyz.com"):
            email_message = email.message_from_string(raw_email)
            for part in email_message.walk():               
                body = part.get_payload()
                bodyList.append(base64.decodestring(body))

        if(sender_of_interest=="wholesalesupplier@abc.com"):
            raw_email_string = raw_email.decode('utf-8',errors='ignore')
            email_message = email.message_from_string(raw_email_string)
            for part in email_message.walk():               
                if part.get_content_type() == "text/plain":
                    body = part.get_payload()
                    body.decode('utf-8',errors='ignore')
                    bodyList.append(body)
    parseBody(bodyList)
    
def parseBody(bodyList):
    for emailBody in bodyList:
        workOrder = ""
        appliance = ""
        warrantyProvider = ""
        address = ""    
        bodyDump = ""
        searchLen = 0

        # We will search the email body for keywords, and depending on that, do a certain process (we have identified that provider now).
        # Another fun thing we will see is parsing strings, stripping them, getting substrings, and seting them to various values.
        
        woSearch2 = -1
        woSearch2 = emailBody.find("Keywords1")
        if(woSearch2>-1):
            warrantyProvider = "Provider1"         
            addrSearch1 = -1
            addrSearch1 = emailBody.find("Address:")
            if(addrSearch1>-1):
                addrSearch1a = emailBody.find(" ",addrSearch1)
                addrSearch1b = emailBody.find("\n",addrSearch1a)
                address = emailBody[addrSearch1a:addrSearch1b].strip()
                
            bodySearch1 = -1
            bodySearch1 = emailBody.find("Service Administrator:")
            searchLen = len("Service Administrator:")
            if(bodySearch1>-1):
                bodySearch1a = emailBody.find("Item Cap/Limit",bodySearch1+searchLen)
                bodyDump = emailBody[bodySearch1:bodySearch1a]
        
            loc1 = emailBody.find("ID", woSearch2)
            loc2 = emailBody.find(" ",loc1) + 1
            loc3 = emailBody.find("\n",loc2)
            workOrder = emailBody[loc2:loc3].rstrip()

        #Provider2
        woSearch1 = -1
        woSearch1 = emailBody.find("Provider2")
        searchLen = len("Provider2")
        if(woSearch1>-1):
            #if Provider2 string has appeared, we know this is Provider2 
            warrantyProvider = "CHW"
            woSearch1a = emailBody.find("#",woSearch1) + 1
            woSearch1b = emailBody.find("\n",woSearch1a + 3)   #let's give a cushion of 3 char's
            workOrder = emailBody[woSearch1a:woSearch1b]
            
            #Provider2: Appliance
            applSearch1 = -1
            applSearch1 = emailBody.find("Reason For Call:")
            searchLen = len("Reason For Call:")
            if(applSearch1>-1):
                applSearch1a = emailBody.find(" ",applSearch1+searchLen)
                applSearch1b = emailBody.find("\n",applSearch1a + 3)   #let's give a cushion of 3 char's
                appliance = emailBody[applSearch1a:applSearch1b]
        
            #Provider2: Address
            addrSearch1 = emailBody.find("Customer:")
            searchLen = len("Customer:")
            if(addrSearch1>-1):
                addrSearch1a = emailBody.find("\n ",addrSearch1+searchLen + 1)
                addrSearch1b = emailBody.find("\n",addrSearch1a+4)
                addrSearch1c = emailBody.find("\n",addrSearch1b+4)
                addrSearch1d = emailBody.find("\n",addrSearch1c+4)
                address = emailBody[addrSearch1b:addrSearch1d]

            #Provider2: Email Body
            bodySearch1 = -1
            bodySearch1 = emailBody.find("Keywords 2")
            searchLen = len("Keywords 2")
            if(bodySearch1>-1):
                bodySearch1a = emailBody.find("You are responsible to collect the",bodySearch1+searchLen)
                bodyDump = emailBody[bodySearch1:bodySearch1a]

        # We have now gathered necessary information from the emails (or from a website)
        # Let's use this text to create our Google Calendar events:

        title1 = ""
        workOrder = str(workOrder).strip()
        warrantyProvider = str(warrantyProvider).strip()
        appliance = str(appliance).strip()

        title1 = workOrder + "-" + appliance + "-" + warrantyProvider
        #remember - checking for dupes is built into createGoogleCalEvent function
        createGoogleCalEvent(title1, address.strip(),bodyDump.strip(), warrantyProvider)
                
def createGoogleCalEvent(title, address, bodyDump, warrantyProvider):    
    jobName = ""
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    #We have to check for duplicates before creating a new "job" on Google Calendar!

    now = datetime.datetime.utcnow().isoformat() + 'Z'
    events_result = service.events().list(calendarId='primary', singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items',[])
    if(warrantyProvider=="Provider1"):
        firstHyph = title.find("-",1)
        secondHyph = title.find("-",firstHyph+1)
        jobName = title[:secondHyph]
    for event in events:
        if ((event['summary']).find(jobName)>-1):
            print ("BusinessXYZScan: found duplicate for following job and continuing on: "+jobName)
            return

    #if we have gotten to this point, means no dup has been found. we can go forward with creating the
    #event for noon of the upcoming Sunday!
    retVals = []
    retVals = createGoogleSheet(title)
    sheetLink = retVals[0]
    
    bodyDump = sheetLink + "\n" + bodyDump

    event = {
    'summary' : '' + title + '',
    'location': '' + address + '',
    'description': '' + bodyDump + '',
    'start': {
    'dateTime': '' + nextSunday()[0] + '',
    'timeZone': 'America/New_York',
    },
    'end': {
    'dateTime': '' + nextSunday()[1] +'',
    'timeZone': 'America/New_York',
    },
    'reminders': {
     'useDefault': False,
     'overrides': [
       {'method': 'email', 'minutes': 24 * 60},
       {'method': 'popup', 'minutes': 10},
     ],
    },
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print ("BusinessXYZScan: an event has been added to your Google Calendar.")

def clear(): 
    # for windows 
    if name == 'nt': 
        _ = system('cls') 
    # for mac and linux(here, os.name is 'posix') 
    else: 
        _ = system('clear') 

def nextSunday():
    retVal = []
    today = datetime.date.today()
    nextSunday = (today + datetime.timedelta( (6-today.weekday()) % 7 ))
    t = datetime.time(hour=12, minute=00)
    t2 = datetime.time(hour=13,minute=00)
    nextSundayStart = (datetime.datetime.combine(nextSunday,t)).isoformat("T")
    nextSundayEnd = (datetime.datetime.combine(nextSunday,t2)).isoformat("T")
    #return [nextSundayStart,nextSundayEnd]
    retVal.append(nextSundayStart)
    retVal.append(nextSundayEnd)
    return retVal

def main():
    clear()
    print("*** BusinessXYZScan.py VERSION v1.0 ***")
    
    #We will read for all emails coming into our new gmail account, from urosmilojkovic. Those emails should contain "job id's".
    #Once we have parsed that data, we will use the information in these emails to create Google Calendar events to help organize our
    #theoretical company.

    readEmail("BusinessXYZ@gmail.com","password123","urosmilojkovic@BusinessXYZ.com")
    print ("BusinessXYZScan has completed execution! See you next time.")
    
if __name__== "__main__":
    main()
