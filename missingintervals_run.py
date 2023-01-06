#!/usr/bin/env python
# coding: utf-8

# In[18]:


import re
import datetime
from time import sleep
from datetime import datetime
from datetime import date
from datetime import time
from datetime import timedelta
from enum import Enum

import os
import fnmatch
import win32com.client as win32
import dateutil
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta




# =====================================BEGINNING OF EMAIL RETRIEVAL=========================================================

# In[2]:


class OutlookFolder(Enum):
    olFolderDeletedItems = 3 # The Deleted Items folder
    olFolderOutbox = 4 # The Outbox folder
    olFolderSentMail = 5 # The Sent Mail folder
    olFolderInbox = 6 # The Inbox folder
    olFolderDrafts = 16 # The Drafts folder
    olFolderJunk = 23 # The Junk E-Mail folder


# In[3]:

print('Created by Joshua McMahon for Lumin8 Transportation Technologies, JAN 2022')
sleep(2)

def stage_program():
    
    user_actual = os.getlogin()
    path = f'C:\\Users\\{user_actual}\\Documents\\Missing Intervals\\'
    
    return path, user_actual


# In[4]:



#IDENTIFY USER
user_actual = os.getlogin()
path = f'C:\\Users\\{user_actual}\\Documents\\Missing Intervals\\'


# In[5]:




# In[6]:


#FETCH ALL DAILY REPORTS FROM EMAIL AND DOWNLOAD TO COMPUTER


def aquire_logs():
    # get a reference to Outlook
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # get the Inbox folder (you can a list of all of the possible settings at https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders)
    inbox = outlook.GetDefaultFolder(OutlookFolder.olFolderInbox.value)
    # get subfolder of this
    missing_intervals = inbox.Folders.Item("Missing Intervals")
    # get all the messages in this folder

    messages = missing_intervals.Items
    # check messages exist
    if len(messages) == 0:
        print("There aren't any messages in this folder")
        exit()

    # loop over them all
    emails = []
    for message in messages:
        # get some information about each message in a tuple

        this_message = (
        message.Subject,
        message.SenderEmailAddress,
        message.To,
        message.Unread,
        message.Senton.date(),
        message.body,
        message.Attachments
        )

        # add this tuple of info to a list holding the messages
        emails.append(this_message)
    # show the result
    for email in emails:
        # unpack the tuple to get at information
        subject, from_address, to_address, if_read, date_sent, body, attachments = email

        # show the subject
        print(subject, date_sent, from_address)
    
        stamp = str(date_sent)
        # number of attachments
        if len(attachments) == 0:
            print("No attachments")
    
        else:
            for attachment in attachments:
                attachment.SaveAsFile(f"C:\\Users\\{user_actual}\\Documents\\Missing Intervals\\" +stamp+ ' ' +attachment.FileName)
                print("Saved {0} attachments".format(len(attachments)))
                


# =================BEGINNING OF TEXT SELECTION==================

# In[7]:


#TODAY'S DATE FOR COMPARISON=========================
def todayis():
    u = datetime.today()
    u = (u.year, u.month, u.day)
    u = datetime(u[0],u[1],u[2])
    u = u.date()
    return u


# In[8]:


#SEVEN DAYS AGO FOR COMPARISON
def lastweek():
    weekago = timedelta(weeks=1)
    lastweek = todayis() - weekago
    return lastweek


# In[9]:


#TIME INTERVAL SELECTION

def fetch_today():
    
    mi_list = []
    for file in os.listdir(path):
        date_pattern = '\d{4}-\d{2}-\d{2}'
        date_match = re.findall(date_pattern, file)
    
        target_today = todayis().isoformat()   
        #target_test = "2022-01-04"
        #target_week = 0  
        pattern = f'*{target_today}*.txt'   
        
        if fnmatch.fnmatch(file,pattern):     #find matching dates then append to a list
            mi_list.append(file)
            print(file)
    return mi_list       
        


# In[10]:


#COMBINE SEARCHED FILES INTO ONE MASTER REQUEST FILE ==============================
def total_request(mi_list):
    with open( path+'0_TOTAL_OUTPUT_REQUEST.txt', 'w') as outfile:
        for fname in mi_list:          #retrieve list of files and then compress into a singular file
            with open(path+fname) as infile:
                outfile.write(infile.read())
                gap ='\n\n\n\n       %%%%%%%%%%%%%%%% NEXT REPORT  %%%%%%%%%%%%%% \n\n\n\n\n  '
                outfile.write(gap)
                print('Log added to request file')
        return


# In[11]:


def sensors(daily):
    sensor_pattern = "\d{6}-\d{1,}"
    sensor_match = re.findall(sensor_pattern,daily) 
    
    return sensor_match


# In[12]:


def coordinates(daily):
    gps_pattern = "(\d{2}.\d{5,6})\s(-\d{2}.\d{5,6})"
    gps_match = re.findall(gps_pattern, daily)
    #for i in gps_match:
     #   print(i)
    return gps_match
    


# In[13]:


def downtime(daily):
    time_pattern = "\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2}:\d{2} \D\D"
    time_match = re.findall(time_pattern,daily)
    print('Start: ' + time_match[0])
    print('End: '+time_match[len(time_match)-1] )
    delta = relativedelta(parse(time_match[0]),parse(time_match[len(time_match)-1])) 
    
    
    return delta


# In[14]:


def report(daily, x):
    daily = daily
    for i in x:
        if (i == x[0]):
            continue
        print("Site Info: ")
        print(sensors(i))
        print('\n')
        print("Coordinates: " )
        print(coordinates(i))
        print('\n')
        print("Length of missed intervals: ")
        print(downtime(i))
        print('\n')
        print('*****************')
    return


# In[15]:



def read_output():
    #SELECT THE TEXT FILE FOR DAILY MISSING INTERVALS
    #text = open('C:\\Users\\Joshua McMahon\\Documents\\Missing Intervals\\Daily_Missing_Intervals_DATE[218].txt')
    text = open(path+'0_TOTAL_OUTPUT_REQUEST.txt')
    daily = text.read()

    #DAILY INTERVAL SPLIT PROCESSING
    x = daily.split("##############################################################################################################################")
    return daily, x


# In[16]:


def run_program():
    print('Begin program')
    sleep(2)
    print('Begin staging')
    sleep(0.1)
    stage_program()
    print('Staging Complete')
    sleep(0.2)
    print('Begin log aquisition')
    sleep(.1)
    aquire_logs()
    sleep(1)
    print('\nFetching...\n')
    mi_list = fetch_today()
    sleep(.2)
    print('Compressing logs')
    total_request(mi_list)
    print('Compression complete')
    sleep(.1)
    daily, x = read_output()
    print('\nGenerating report....\n')
    sleep(2)
    print('\n=== Output written -------> \n\n\n')
    report(daily, x)
    print('End of Report')
    
    
    
    input('Press "ENTER" to quit program')
    
    
    


# In[17]:


run_program()


# In[ ]:




