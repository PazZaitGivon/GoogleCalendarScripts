from __future__ import print_function
import httplib2
import os
import datetime
import json

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_id.json'#replace this with a file you get from google
APPLICATION_NAME = 'Croudsourcable Google Calendar'
with open('SheetAndCalendarId.json') as json_file: #replace this with a specially formatted JSON file, I will post an example
  IDENTIFIERS = json.load(json_file)
  print (IDENTIFIERS)
  sheets=IDENTIFIERS["sheetIds"]
  formName0=sheets[0]["name"]
  sheetId0=sheets[0]["id"]

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    """Shows basic usage of the Sheets API."""
    #Plan: scroll from last to first entry (can we do that?) till we reach a 100 or the first entry. 
    #as we processes entries if 0 turn to 100, if 1 create event and turn to 100 if blank ignore 
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = sheetId0
    rangeName = formName0
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:

        rowIndex=len(values)-1
        print("the last row is: "+`rowIndex`)#DEBUG
        #1-EventTitle   2- EventType  3-NONEXISTANT(all day event?) 4-StartTime 
        #5-Adress 6-Summary 7-Status(not autopopulated) 8-Endtime 9-Date
        while rowIndex>0:
            print("On row "+`rowIndex`)#DEBUG
            # if 1 (we want to approve this) print the whole thing
            if((values[rowIndex][7])==""):
              print("still undecided on "+`rowIndex`)
            elif(int(`values[rowIndex][7]`[2:-1])==1):
                print(`values[rowIndex]`)
                date=values[rowIndex][9]
                print(date)#we may need to zero pad month and day
                if(date[1]=='/'):
                  print("zero padding month")
                  date='0'+date
                  print(date)
                if(date[4]=='/'):
                  print("zero padding day")
                  date=date[0:3]+'0'+date[3:]
                  print(date)
                #What we really want to do here is make a calendar statement, switch to -100
            elif(int(`values[rowIndex][7]`[2:-1])==0):#this event was rejected
                print("was zero")#DEBUG
                #change it to -200 Things to think about, maybe we could reject with a string explanation of why? so anything other than 1 is reject?
            rowIndex=rowIndex-1

def eventCreator(title,startTime, endTime, location,description):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    event = {
      'summary': title,
      'location': location,
      'description': description,
      'start': {
        'dateTime': startTime, #'2017-02-28T09:00:00-07:00',
        #'timeZone': 'America/Los_Angeles',
      },
      'end': {
        'dateTime': endTime, #'2017-02-28T09:20:00-07:00',
        #'timeZone': 'America/Los_Angeles',
      },
      #'recurrence': [
      #  'RRULE:FREQ=DAILY;COUNT=2'
      #],

    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print ('Event created:'+`(event.get('htmlLink'))` )



if __name__ == '__main__':
    main()
    #eventCreator('Test2','2017-03-04T09:00:00-07:00','2017-03-04T09:20:00-07:00','800 Howard St., San Francisco, CA 94103',"A chance to hear more about Google\'s developer products.")
