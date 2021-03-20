import requests
import os
from datetime import date, datetime, timedelta as td
import pandas as pd

#For Google Sheets
from oauth2client.service_account import ServiceAccountCredentials
import gspread



#Credentials
import json

with open("creds.json", "r") as file:
    credentials = json.load(file)
    KEY = credentials['rescuetime_KEY']

baseurl = 'https://www.rescuetime.com/anapi/data?key='
url =  baseurl + KEY


## Export Dates Configuration
# Configure These to Your Preferred Dates
start_date = '2020-12-19'  # Start date for additional data export
end_date   = '2020-12-22'  # End date for data

# Adjustable by Time Period
def rescuetime_get_activities(start_date, end_date, sheet, resolution='hour'):


    #Find last added date and current date
    lastRowIndex = len(sheet.col_values(1))
    lastRow = sheet.row_values(lastRowIndex)
    lastRowDate = lastRow[0][0:10]
    yesterday = (date.today()-td(days=1)).strftime("%Y-%m-%d")

    # Configuration for Query
    # SEE: https://www.rescuetime.com/apidoc
    payload = {
        'perspective':'interval',
        'resolution_time': resolution, #1 of "month", "week", "day", "hour", "minute"
        'restrict_kind':'document',
        'restrict_begin': lastRowDate,
        'restrict_end': yesterday,
        'format':'json' #csv
    }
    
    # Setup Iteration - by Day
    d1 = datetime.strptime(payload['restrict_begin'], "%Y-%m-%d").date()
    d2 = datetime.strptime(payload['restrict_end'], "%Y-%m-%d").date()
    delta = d2 - d1
    
    activities_list = []
    
    # Iterate through the days, making a request per day
    for i in range(delta.days + 1):
        # Find iter date and set begin and end values to this to extract at once.
        d3 = d1 + td(days=i) # Add a day
        if d3.day == 1: print('Pulling Monthly Data for ', d3)

        # Update the Payload
        payload['restrict_begin'] = str(d3) # Set payload days to current
        payload['restrict_end'] = str(d3)   # Set payload days to current

        # Request
        try: 
            r = requests.get(url, payload) # Make Request
            iter_result = r.json() # Parse result
            # print("Collecting Activities for " + str(d3))
        except: 
            print("Error collecting data for " + str(d3))
    
        if len(iter_result) != 0:
            for i in iter_result['rows']:
                activities_list.append(i)
        else:
            print("Appears there is no RescueTime data for " + str(d3))

    #Reorder list
    order = [0, 3, 4, 0, 5, 1]    

    result = [] 
    for a in activities_list:
        result.append([a[i] for i in order])


    return result

def authorize_sheets():
    scopesGoogleSheets = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds_sheets.json", scopesGoogleSheets)
    client = gspread.authorize(creds)
    sheet = client.open("RescueTime new").sheet1  # Open the spreadhseet

    return sheet

def insert_into_sheets(activities_list, sheet):
    for activity in activities_list:
        #print("Activity 0: "+ str(activity[0]))
        #print("Activity 1: "+ str(activity[1]))
        #print("Activity 2: "+ str(activity[2]))
        #print("Activity 3: "+ str(activity[3]))
        #print("Activity 4: ".encode('utf-8')+ str(activity[4]).encode('utf-8'))
        #sheet.insert_row([activity[0], activity[1]])
    sheet.insert_rows(activities_list)


def rescuetime_get_activities_daily_summaries(start_date, end_date, resolution='hour'):
    #Daily summaries
    baseurl = "https://www.rescuetime.com/anapi/daily_summary_feed?key="
    url = baseurl + KEY
    # Request
    try: 
        r = requests.get(url) # Make Request
        iter_result = r.json() # Parse result
        # print("Collecting Activities for " + str(d3))
    except: 
        print("Error collecting data for " + str(d3))

    index = 0
    for result in iter_result:
        print(result)

        print("index is = " + str(index))
        index+=1





# Defining main function 
def main(): 
    sheet = authorize_sheets()
    activities_list = rescuetime_get_activities(start_date, end_date, sheet)

    #for activity in activities_list:
        #print(activity)

    insert_into_sheets(activities_list, sheet)
  
  
# Using the special variable  
# __name__ 
if __name__=="__main__": 
    main() 