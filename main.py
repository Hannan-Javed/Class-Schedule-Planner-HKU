import pandas as pd
import datetime as dt
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

excel_file = '2024-25_class_timetable_20240722.xlsx'
SCOPES = ["https://www.googleapis.com/auth/calendar"]
pd.set_option('display.max_columns', None)

def getCourseList():

    df = pd.read_excel(excel_file, skiprows=None)
    df.columns = df.columns.str[1:]
    
    degree = input("Please enter your degree (UG/TPG/RPG): ")
    while degree.upper() not in ["UG", "TPG", "RPG"]:
        degree = input("Invalid input. Please enter 'UG', 'TPG', or 'RPG': ")

    courses = df[df['ACAD_CAREER'].str.contains(degree.upper())]

    return courses

def main():

    def addToCalender(course):

        mode = input("Do you want to add courses for one week (for planning) or for the whole semester? (week/semester) ")

        for index, row in course.iterrows():

            start_date = row['START DATE'].date()
            end_date = row['END DATE'].date()
            start_time = row['START TIME']
            end_time = row['END TIME']
            title = row['COURSE TITLE']
            description = row['VENUE']
            section = row['CLASS SECTION']
            print(start_time, end_time, title, description, section, type(start_time), type(end_time), type(title), type(description), type(section))

            event = {
                'summary': f'{title} {section}',
                'description': f'{description}',
                'start': {
                    'dateTime': f'{start_date}T{start_time}+08:00',
                    'timeZone': 'Asia/Hong_Kong',
                },
                'colorId': 7,
                'end': {
                    'dateTime': f'{start_date}T{end_time}+08:00',
                    'timeZone': 'Asia/Hong_Kong',
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                    ],
                },
                
            }
            if mode == 'semester':
                days = ','.join([day[:2] for day in ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'] if pd.notnull(row[day])])
                event['recurrence'] = [f'RRULE:FREQ=WEEKLY;UNTIL={end_date};BYDAY={days}',]
        
            # Insert event to calendar
            event = service.events().insert(calendarId='primary', body=event).execute()
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    try:
        service = build('calendar', 'v3', credentials=creds)
        now = dt.datetime.utcnow().isoformat() + 'Z'
        events_result = service.events().list(calendarId='primary', timeMin=now, maxResults=10, singleEvents=True, orderBy='startTime').execute()
        print(events_result)

        
            
        print("Welcome to HKU Course Planner!")
        print("At any point in the planning if you want to up one level in the menu, please enter '-1'.")

        course_list = getCourseList()

        # main program
        semester = '0'
        search_mode = '0'
        add_more = 'y'
        while semester != '-1':
            semester = input("Please enter the semester: (sem1/sem2) ")
            while semester[-1] != '1' and semester[-1] != '2' and semester != '-1':
                semester = input("Wrong input, Please enter 'sem1' or 'sem2': ")
            semester = "Sem 1" if semester[-1] == '1' else "Sem 2"
            course_list = course_list[course_list['ERM'].str.contains(semester)]
            while search_mode != '-1':
                search_mode = input("Search course by course code or course name? (code/name) ")
                while search_mode[0].lower() != 'c' and search_mode[0].lower() != 'n' and search_mode != '-1':
                    search_mode = input("Wrong input, Please enter 'code' or 'name': ")

                while add_more[0].lower() == 'y':
                    search = input(f"Please enter the course {"code" if search_mode[0].lower() == 'c' else "title"}: ")
                    search_result = course_list[course_list["COURSE CODE" if search_mode[0].lower() == 'c' else "COURSE TITLE"].str.contains(search.upper())]
                    if search_result.empty:
                        print("No course found.")
                    else:
                        print(search_result)
                        addcourse = input("Do you want to add the course to your google calender? (y/n) ")
                        if addcourse[0].lower() == 'y':
                            addToCalender(search_result)
                        else:
                            add_more = input("Do you want to search for more courses? (y/n) ")

        

    except HttpError as e:
        print("An error occured: ", e)



if __name__ == '__main__':
    main()