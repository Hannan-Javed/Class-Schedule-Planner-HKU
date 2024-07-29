import pandas as pd
import datetime as dt
from datetime import timedelta
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
mode_dict = {1: 'week', 2: 'semester'}
semester_dict = {1: 'Sem 1', 2: 'Sem 2'}
search_mode_dict = {1: 'COURSE CODE', 2: 'COURSE TITLE'}
courses_added = {}
day_index = {'MON': 0, 'TUE': 1, 'WED': 2, 'THU': 3, 'FRI': 4}


def main():

    def addToCalender(course, testmode):


        if course['CLASS SECTION'].nunique() > 1:
            add_sections = input("There are more than one section available for the chosen courses in this semester. Do you want to add all sections of the course? (y/n) ")
            while add_sections[0].lower() not in ['y', 'n']:
                add_sections = input("Invalid input. Please enter y or n: ")
            if add_sections[0].lower() == 'n':
                add_sections = input(f"{','.join(course['CLASS SECTION'].unique())}\nPlease enter the section(s) you want to add. If more than one section, then seperate with comma: ").upper()
                while not all(section in course['CLASS SECTION'].unique() for section in add_sections.split(',')):
                    add_sections = input(f"{','.join(course['CLASS SECTION'].unique())}\nInvalid section. Please enter a valid section: ").upper()
        else:
            add_sections = 'y'
        
        add_section_title = input("Do you want to add the section title together with the course title? (y/n) ")
        while add_section_title[0].lower() not in ['y', 'n']:
            add_section_title = input("Invalid input. Please enter y or n: ")

        for index, row in course.iterrows():

            if mode_dict[int(testmode)] == 'week' and (row['COURSE CODE'] + ' ' + row['CLASS SECTION'] + ' ' + str([day_index[row[day]] for day in day_index.keys() if pd.notnull(row[day])][0])) in courses_added.values():
                continue

            if add_sections == 'y' or row['CLASS SECTION'] in add_sections.upper().split(','):

                start_date = row['START DATE'].date()
                end_date = row['END DATE'].date()
                start_time = row['START TIME']
                end_time = row['END TIME']
                code = row['COURSE CODE']
                title = row['COURSE TITLE OG']
                description = row['VENUE']
                section = row['CLASS SECTION']
                days_difference = [day_index[row[day]] for day in day_index.keys() if pd.notnull(row[day])][0]

                if mode_dict[int(testmode)] == 'week':
                    today = pd.Timestamp('today').date()
                    if 'Sem 1' in row['ERM']:
                        start_date = today + timedelta(days=(7 - today.weekday()))  # Next Monday
                    elif 'Sem 2' in row['ERM']:
                        start_date = today + timedelta(days=(14 - today.weekday()))  # Monday after next

                    start_date = start_date + timedelta(days=days_difference)
                    
                event = {
                    'summary': f'{code} {title} {section if add_section_title[0].lower() == 'y' else ''}',
                    'description': f'{description}',
                    'start': {
                        'dateTime': f'{start_date}T{start_time}+08:00',
                        'timeZone': 'Asia/Hong_Kong',
                    },
                    # default color is blue
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
                if mode_dict[int(testmode)] == 'semester':
                    event['recurrence'] = [f'RRULE:FREQ=WEEKLY;UNTIL={(int(str(end_date).replace('-',''))+1)};BYDAY={[day[:2] for day in day_index.keys() if pd.notnull(row[day])][0]}',]

                # Insert event to calendar
                print(f"Adding {row['COURSE CODE']} {row['CLASS SECTION']} from {start_date} to {end_date if mode_dict[int(testmode)] == 'semester' else start_date} ({[key for key, value in day_index.items() if value == days_difference][0]}) to your google calendar...")
                event = service.events().insert(calendarId='primary', body=event).execute()

                courses_added[event['id']] = code + ' ' + section + ' ' + str(days_difference)
        


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
        
        while True:
            degree = input("\n1. UG - Undergraduate\n2. TPG - Taught Postgraduate\n3. RPG - Research Postgraduate\n4. Exit\nPlease enter your degree: ")

            while not degree.isdigit() or int(degree) not in range(1, 5):
                degree = input("\n1. UG - Undergraduate\n2. TPG - Taught Postgraduate\n3. RPG - Research Postgraduate\n4. Exit\nInvalid input. Please select an option again: ")

            if degree == '4':
                break

            while degree != '4':

                df = pd.read_excel(excel_file, skiprows=None)
                df.columns = df.columns.str[1:]
                df['COURSE TITLE OG'] = df['COURSE TITLE']
                df['COURSE TITLE'] = df['COURSE TITLE'].str.lower()
                course_list = df[df['ACAD_CAREER'].str.contains(degree_dict[int(degree)])]
            

                mode = input("\n1. One week\n2. Whole semester\n3. Go back\nDo you want to add courses for one week (for planning) or for the whole semester? ")
                while not mode.isdigit() or int(mode) not in range(1, 4):
                    mode = input("\n1. One week\n2. Whole semester\n3. Go back\nInvalid input. Please select an option again: ")

                if mode == '1':
                    print("In this mode, the semester 1 courses will be added to the nearest next week, and the semester 2 courses will be added to the week after the semester 1 courses.")
                if mode == '3':
                    break                

                while mode != '3':

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

                            search_result = course_list[course_list[search_mode_dict[int(search_mode)]].str.contains(search.lower())]
                            if search_result.empty:
                                print("No course found.")
                            else:
                                print(search_result)
                                addcourse = input("Do you want to add the course to your google calendar? (y/n) ")

                                while addcourse[0].lower() not in ['y', 'n']:
                                    addcourse = input("Invalid input. Please enter y or n: ")

                                if addcourse[0].lower() == 'y':
                                    if search_mode == '2' and search_result['COURSE CODE'].nunique() > 1:

                                        for i, (code, title) in enumerate(zip(search_result['COURSE CODE'].unique(), search_result['COURSE TITLE'].unique()), 1):
                                            print(f"{i}. {code} - {title}")

                                        courses = input("Please enter the numbers of the courses you want to add (separated by commas): ")
                                        while not all(course.isdigit() and int(course) <= search_result['COURSE TITLE'].nunique() for course in courses.split(',')):
                                            courses = input(f"Invalid input. Please enter valid course number(s): ")

                                        for courses in courses.split(','):
                                            addToCalender(search_result[search_result['COURSE TITLE'] == search_result['COURSE TITLE'].unique()[int(courses)-1]], mode)    
                                    else:                                  
                                        addToCalender(search_result, mode)

                            add_more = input(f"Do you want to search for more courses by {search_mode_dict[int(search_mode)].lower()}? (y/n) ")
                            while add_more[0].lower() not in ['y', 'n']:
                                add_more = input("Invalid input. Please enter y or n: ")
                            if add_more[0].lower() == 'n':
                                break
                    if len(courses_added) > 0:
                        clear = input("Do you want to clear all the entered events? (y/n) ")
                        while clear[0].lower() not in ['y', 'n']:
                            clear = input("Invalid input. Please enter y or n: ")

                        if clear[0].lower() == 'y':
                            for course_id in courses_added.keys():
                                print(f"Deleting {courses_added[course_id][:8]} {courses_added[course_id][9:11]} {[day for day,dif in day_index.items() if dif==int(courses_added[course_id][-1])][0]} from your google calendar...")
                                service.events().delete(calendarId='primary', eventId=course_id).execute()
                            courses_added.clear()

        print("Thank you for using HKU Course Planner!")
    except HttpError as e:
        print("An error occured: ", e)



if __name__ == '__main__':
    main()