import pandas as pd
import datetime as dt
import os.path, re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# global variables
excel_file = '2024-25_class_timetable_20240722.xlsx'
SCOPES = ["https://www.googleapis.com/auth/calendar"]
pd.set_option('display.max_columns', None)
degree_dict = {1: 'UG', 2: 'TPG', 3: 'RPG'}
semester_dict = {1: 'Sem 1', 2: 'Sem 2'}
search_mode_dict = {1: 'COURSE CODE', 2: 'COURSE TITLE'}


def main():

    def addToCalender(course):

        for index, row in course.iterrows():

            start_date = row['START DATE'].date()
            end_date = row['END DATE'].date()
            start_time = row['START TIME']
            end_time = row['END TIME']
            title = row['COURSE TITLE']
            description = row['VENUE']
            section = row['CLASS SECTION']

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
            else:
                break
        
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

    
        # main program  
        print("Welcome to HKU Course Planner!")

        df = pd.read_excel(excel_file, skiprows=None)
        df.columns = df.columns.str[1:]
        
        while True:
            degree = input("\n1. UG - Undergraduate\n2. TPG - Taught Postgraduate\n3. RPG - Research Postgraduate\n4. Exit\nPlease enter your degree: ")

            while not degree.isdigit() or int(degree) not in range(1, 5):
                degree = input("\n1. UG - Undergraduate\n2. TPG - Taught Postgraduate\n3. RPG - Research Postgraduate\n4. Exit\nInvalid input. Please select an option again: ")

            if degree == '4':
                break

            while degree != '4':
                course_list = df[df['ACAD_CAREER'].str.contains(degree_dict[int(degree)])]

                mode = input("\n1. One week\n2. Whole semester\n3. Go back\nDo you want to add courses for one week (for planning) or for the whole semester? ")
                while not mode.isdigit() or int(mode) not in range(1, 4):
                    mode = input("\n1. One week\n2. Whole semester\n3. Go back\nInvalid input. Please select an option again: ")

                if mode == '3':
                    break

                while mode != '3':
                    if mode == '1':
                        print("In this week, the semester 1 courses will be added to the nearest next week, and the semester 2 courses will be added to the week after the semester 1 courses.")

                    semester = input("\n1. Sem 1\n2. Sem 2\n3. Go back\nPlease enter the semester: ")
                    while not semester.isdigit() or int(semester) not in range(1, 4):
                        semester = input("\n1. Sem 1\n2. Sem 2\n3. Go back\nInvalid input. Please select an option again: ")

                    if semester == '3':
                        break

                    while semester != '3':
                        course_list = course_list[course_list['ERM'].str.contains(semester_dict[int(semester)])]

                        search_mode = input("\n1. Search by course code\n2. Search by course title\n3. Go back\nPlease enter searching field: ")
                        while not search_mode.isdigit() or int(search_mode) not in range(1, 4):
                            search_mode = input("\n1. Search by course code\n2. Search by course title\n3. Go back\nInvalid input. Please select an option again: ")

                        if search_mode == '3':
                            break

                        search = ''
                        while search_mode != '3' and search != '-1':
                            search = input(f"Please enter the {search_mode_dict[int(search_mode)].lower()}: (-1 to go back) ")
                            

                            while search_mode == '1' and not re.match("[a-zA-Z]{4}[0-9]{4}", search) and search != '-1':
                                search = input(f"Invalid course code format. Please enter the {search_mode_dict[int(search_mode)].lower()} again: (format is XXXXDDDD) ")

                            if search == '-1':
                                break

                            search_result = course_list[course_list[search_mode_dict[int(search_mode)]].str.contains(search.upper())]
                            if search_result.empty:
                                print("No course found.")
                            else:
                                print(search_result)
                                addcourse = input("Do you want to add the course to your google calendar? (y/n) ")

                                while addcourse[0].lower() not in ['y', 'n']:
                                    addcourse = input("Invalid input. Please enter y or n: ")

                                if addcourse[0].lower() == 'y':
                                    addToCalender(search_result)

                            add_more = input("Do you want to search for more courses? (y/n) ")
                            while add_more[0].lower() not in ['y', 'n']:
                                add_more = input("Invalid input. Please enter y or n: ")
                            if add_more[0].lower() == 'n':
                                break

        print("Thank you for using HKU Course Planner!")
    except HttpError as e:
        print("An error occured: ", e)



if __name__ == '__main__':
    main()