import pandas as pd
import datetime as dt
from datetime import timedelta
import os.path, re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from PyInquirer import prompt
# global variables
excel_file = '2024-25_class_timetable_20240722.xlsx'
SCOPES = ["https://www.googleapis.com/auth/calendar"]
pd.set_option('display.max_columns', None)
degree_dict = {"UG - Undergraduate": 'UG', "TPG - Taught Postgraduate": 'TPG', "RPG - Research Postgraduate": 'RPG'}
search_mode_dict = {"Course code": 'COURSE CODE', "Course title": 'COURSE TITLE'}
courses_added = {}
day_index = {'MON': 0, 'TUE': 1, 'WED': 2, 'THU': 3, 'FRI': 4}

def get_default_download_directory():
    home_directory = os.path.expanduser("~")  # Get user's home directory
    
    # Check the operating system to determine the download directory
    if os.name == 'posix':  # Linux or macOS
        download_directory = os.path.join(home_directory, 'Downloads')
    elif os.name == 'nt':  # Windows
        download_directory = os.path.join(home_directory, 'Downloads')
    else:
        download_directory = None  # Unsupported operating system
    return download_directory


def main():

    def listMenuSelector(field, qprompt, questions):
        menu = prompt(
                [
                    {
                        'type': 'list',
                        'name': field,
                        'message': qprompt,
                        'choices': questions,
                    }
                ]
            )
        return menu[field]

    def addToCalender(course, testmode):

        if course['CLASS SECTION'].nunique() > 1:
            add_sections = input(f"There are more than one section of {course['COURSE TITLE OG'].unique()[0]}. Please enter the section(s) you want to add. Select multiple seperated by a comma {','.join(course['CLASS SECTION'].unique())}\nSections: ").upper()
            while not all(section in course['CLASS SECTION'].unique() for section in add_sections.split(',')):
                add_sections = input(f"{','.join(course['CLASS SECTION'].unique())}\nInvalid section. Please enter a valid section: ").upper()
        else:
            add_sections = 'y'
        
        add_section_title = listMenuSelector('add_section_title', 'Do you want to add the section title together with the course title?', ['Yes', 'No'])

        for index, row in course.iterrows():

            if testmode == 'One week' and (row['COURSE CODE'] + ' ' + row['CLASS SECTION'] + ' ' + str([day_index[row[day]] for day in day_index.keys() if pd.notnull(row[day])][0])) in courses_added.values():
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

                if testmode == 'One week':
                    today = pd.Timestamp('today').date()
                    if 'Sem 1' in row['ERM']:
                        start_date = today + timedelta(days=(7 - today.weekday()))  # Next Monday
                    elif 'Sem 2' in row['ERM']:
                        start_date = today + timedelta(days=(14 - today.weekday()))  # Monday after next

                    start_date = start_date + timedelta(days=days_difference)
                    
                event = {
                    'summary': f'{code} {title} {section if add_section_title == 'Yes' else ''}',
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
                if testmode == 'Whole semester':
                    event['recurrence'] = [f'RRULE:FREQ=WEEKLY;UNTIL={(int(str(end_date).replace('-',''))+1)};BYDAY={[day[:2] for day in day_index.keys() if pd.notnull(row[day])][0]}',]

                # Insert event to calendar
                print(f"Adding {row['COURSE CODE']} {row['CLASS SECTION']} from {start_date} to {end_date if testmode == 'Whole semester' else start_date} ({[key for key, value in day_index.items() if value == days_difference][0]}) to your google calendar...")
                event = service.events().insert(calendarId='primary', body=event).execute()

                courses_added[event['id']] = code + ' ' + section + ' ' + str(days_difference)
    
    def makeDirectory(term, course):
        choice = listMenuSelector('directory', 'Please select the directory to save the files:', ['Downloads', 'Input Directory Path'])
        
        if choice == 'Downloads':
            directory = get_default_download_directory()
        else:
            directory = input("Please enter the directory path: ")
            if not os.path.exists(directory):
                print("The directory does not exist. Please try again.")
                directory = input("Please enter the directory path: ")

        directory = os.path.join(directory, term, course)
        os.makedirs(directory, exist_ok=True)

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
            degree = listMenuSelector('degree', 'Please select your degree:', ['UG - Undergraduate', 'TPG - Taught Postgraduate', 'RPG - Research Postgraduate', 'Exit'])

            if degree == 'Exit':
                break

            while degree != 'Exit':

                df = pd.read_excel(excel_file, skiprows=None)
                df.columns = df.columns.str[1:]
                df['COURSE TITLE OG'] = df['COURSE TITLE']
                df['COURSE TITLE'] = df['COURSE TITLE'].str.upper()
                course_list = df[df['ACAD_CAREER'].str.contains(degree_dict[degree])]
                
                mode = listMenuSelector('mode', 'Please select if you want to add courses for one week (for planning) or for the whole semester? ', ['One week', 'Whole semester', 'Go back'])
                
                if mode == 'One week':
                    print("In this mode, the semester 1 courses will be added to the nearest next week, and the semester 2 courses will be added to the week after the semester 1 courses.")
                elif mode == 'Go back':
                    break                

                while mode != 'Go back':

                    semester = listMenuSelector('semester', 'Please select the semester:', ['Sem 1', 'Sem 2', 'Go back'])

                    if semester == 'Go back':
                        break

                    while semester != 'Go back':
                        course_list = course_list[course_list['ERM'].str.contains(semester)]

                        search_mode = listMenuSelector('search_mode', 'Please select the searching field:', ['Course code', 'Course title', 'Go back'])
                        if search_mode == 'Go back':
                            break

                        search = ''
                        while search_mode != 'Go back' and search != '-1':
                            search = input(f"Please enter the {search_mode.lower()}: (-1 to go back) ")
                            
                            while search_mode == "Course code" and not re.match("[a-zA-Z]{4}[0-9]{4}", search) and search != '-1':
                                search = input(f"Invalid course code format. Please enter the course code again: (format is XXXXDDDD) ")

                            if search == '-1':
                                break

                            search_result = course_list[course_list[search_mode_dict[search_mode]].str.contains(search.upper())]
                            if search_result.empty:
                                print("No course found.")
                            else:
                                print(search_result)
                                addcourse = listMenuSelector('addcourse', 'Do you want to add the course to your google calendar?', ['Yes', 'No'])

                                if addcourse == 'Yes':
                                    if search_mode == 'Course title' and search_result['COURSE CODE'].nunique() > 1:
                                        for i, (code, title) in enumerate(zip(search_result['COURSE CODE'].unique(), search_result['COURSE TITLE'].unique()), 1):
                                            print(f"{i}. {code} - {title}")

                                        courses = input("Please enter the numbers of the courses you want to add (separated by commas): ")
                                        while not all(course.isdigit() and int(course) <= search_result['COURSE TITLE'].nunique() for course in courses.split(',')):
                                            courses = input(f"Invalid input. Please enter valid course number(s): ")

                                        for courses in courses.split(','):
                                            addToCalender(search_result[search_result['COURSE TITLE'] == search_result['COURSE TITLE'].unique()[int(courses)-1]], mode)    
                                    else:                                  
                                        addToCalender(search_result, mode)

                                makedirectory = listMenuSelector('makedirectory', 'Do you want to make a folder for the courses added?', ['Yes', 'No'])

                                if makedirectory == 'Yes':
                                    course = f"{search_result['COURSE CODE'].unique()[0]} - {search_result['COURSE TITLE OG'].unique()[0]}"
                                    makeDirectory(semester, course)
                            add_more = listMenuSelector('add_more', f'Do you want to search for more courses by {search_mode.lower()}?', ['Yes', 'No'])

                            if add_more == 'No':
                                break

                    if len(courses_added) > 0:
                        clear = listMenuSelector('clear', 'Do you want to clear all the entered events?', ['Yes', 'No'])
                
                        if clear == 'Yes':
                            for course_id in courses_added.keys():
                                print(f"Deleting {courses_added[course_id][:8]} {courses_added[course_id][9:11]} {[day for day,dif in day_index.items() if dif==int(courses_added[course_id][-1])][0]} from your google calendar...")
                                service.events().delete(calendarId='primary', eventId=course_id).execute()
                            courses_added.clear()

        print("Thank you for using HKU Course Planner!")
    except HttpError as e:
        print("An error occured: ", e)



if __name__ == '__main__':
    main()