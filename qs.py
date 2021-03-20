import requests
import os
from datetime import date, datetime, timedelta as td
import pandas as pd
import json

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
def rescuetime_get_activities(start_date, end_date, sheet, resolution='minute'):


    #Find last added date and current date
    lastRowIndex = len(sheet.col_values(1))
    lastRow = sheet.row_values(lastRowIndex)
    lastRowDate = lastRow[0]
    dayAfterLastDate = pd.to_datetime(lastRowDate) + td(days=1)
    dayAfterLastDateString = dayAfterLastDate.strftime("%Y-%m-%d")
    yesterday = (date.today()-td(days=1)).strftime("%Y-%m-%d")

    # Configuration for Query
    # SEE: https://www.rescuetime.com/apidoc
    payload = {
        'perspective':'interval',
        'resolution_time': resolution, #1 of "month", "week", "day", "hour", "minute"
        'restrict_kind':'document',
        'restrict_begin': dayAfterLastDateString,
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
    order = [0, 3, 4, 2, 5, 1]    

    result = [] 
    for a in activities_list:
        result.append([a[i] for i in order])


    return result

def authorize_sheets(sheet_name):
    scopes_google_sheets = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds_sheets.json", scopes_google_sheets)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).sheet1

    return sheet

def insert_into_sheets(list_to_insert, sheet):
    sheet.insert_rows(list_to_insert, row=len(sheet.col_values(1))+1)

def rescuetime_get_activities_daily_summaries(sheet):
    #Daily summaries
    baseurl = "https://www.rescuetime.com/anapi/daily_summary_feed?key="
    url = baseurl + KEY
    
    try: 
        r = requests.get(url) 
        iter_result = r.json()
        print("Collecting daily summaries")
    except: 
        print("Error collecting daily summaries")
    
    results = []
    for result in iter_result:
        results.append(list(result.values()))
    
    #Only add the days that aren't present
    last_row_date = sheet.col_values(2)[-1]


    return results
    

#Used this temporarily to make the date format between the downloaded and from the API consistent.
def makeDateFormatConsistent(sheet):
    dates = sheet.col_values(1)
    incorrectDates = []
    for date in dates:
        if(not "T" in date and not "Date" in date):
            incorrectDates.append(date)

    newDates = []
    for date in incorrectDates:
        date = date[0:19]
        date = datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
        #make this format 2021-02-17T07:30:00
        date = date.strftime("%Y-%m-%dT%H:%M:%S")
        #print(date)
        newDates.append(date)
    nestedArray = []
    for array in newDates:
        nestedArray.append([array])
    sheet.insert_rows(nestedArray, row=2)



def main():
    #Export daily summaries
    sheet = authorize_sheets("RescueTime - Daily summaries")
    daily_summaries = rescuetime_get_activities_daily_summaries(sheet)
    insert_into_sheets(daily_summaries, sheet)

    #Export detailed data
    sheet = authorize_sheets("RescueTime")
    activities_list = rescuetime_get_activities(start_date, end_date, sheet)
    insert_into_sheets(activities_list, sheet)
  
if __name__=="__main__": 
    main() 