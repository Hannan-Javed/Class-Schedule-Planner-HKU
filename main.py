import re
from datetime import timedelta
from PyInquirer import prompt
from excel_reader import ExcelReader
from calendar_manager import CalendarManager
from directory_manager import DirectoryManager
from utils import list_menu_selector
import pandas as pd
# Global variables
EXCEL_FILE = '2024-25_class_timetable_20240722.xlsx'
degree_dict = {"UG - Undergraduate": 'UG', "TPG - Taught Postgraduate": 'TPG', "RPG - Research Postgraduate": 'RPG'}
search_mode_dict = {"Course code": 'COURSE CODE', "Course title": 'COURSE TITLE'}
courses_added = {}
day_index = {'MON': 0, 'TUE': 1, 'WED': 2, 'THU': 3, 'FRI': 4}


def add_to_calendar(calendar_manager, course, testmode):
    if course['CLASS SECTION'].nunique() > 1:
        sections = course['CLASS SECTION'].unique()
        add_sections = input(f"There are multiple sections for {course['COURSE TITLE OG'].iloc[0]}. Enter sections to add (comma-separated): {', '.join(sections)}\nSections: ").upper()
        while not all(section in sections for section in add_sections.split(',')):
            add_sections = input(f"Invalid sections. Please enter valid sections: {', '.join(sections)}\nSections: ").upper()
    else:
        add_sections = 'Y'

    add_section_title = list_menu_selector('add_section_title', 'Add section title to course title?', ['Yes', 'No'])

    for _, row in course.iterrows():
        identifier = f"{row['COURSE CODE']} {row['CLASS SECTION']} {day_index[[day for day in day_index if pd.notnull(row[day])][0]]}"
        if testmode == 'One week' and identifier in courses_added.values():
            continue

        if add_sections == 'Y' or row['CLASS SECTION'] in add_sections.split(','):
            event = create_event(row, testmode, add_section_title)
            print(f"Adding {row['COURSE CODE']} {row['CLASS SECTION']} to Google Calendar...")
            added_event = calendar_manager.add_event(event)
            if added_event:
                courses_added[added_event['id']] = identifier

def create_event(row, testmode, add_section_title):
    start_date = row['START DATE'].date()
    end_date = row['END DATE'].date()
    code = row['COURSE CODE']
    title = row['COURSE TITLE OG']
    section = row['CLASS SECTION']
    description = row['VENUE']
    days_difference = day_index[[day for day in day_index if pd.notnull(row[day])][0]]

    if testmode == 'One week':
        today = pd.Timestamp('today').date()
        if 'Sem 1' in row['TERM']:
            # add to next week
            start_date = today + timedelta(days=(7 - today.weekday()))
        elif 'Sem 2' in row['TERM']:
            # add to next next week
            start_date = today + timedelta(days=(14 - today.weekday()))
        start_date += timedelta(days=days_difference)
    # google calendar event
    event = {
        'summary': f"{code} {title} {' ' + section if add_section_title == 'Yes' else ''}",
        'description': description,
        'start': {
            'dateTime': f"{start_date}T{row['START TIME']}+08:00",
            'timeZone': 'Asia/Hong_Kong',
        },
        'end': {
            'dateTime': f"{start_date}T{row['END TIME']}+08:00",
            'timeZone': 'Asia/Hong_Kong',
        },
        'colorId': 7,
        'reminders': {
            'useDefault': False,
            'overrides': [],
        },
    }

    if testmode == 'Whole semester':
        day = [day[:2] for day in day_index if pd.notnull(row[day])][0]
        until = int(str(end_date).replace('-', '')) + 1
        event['recurrence'] = [f"RRULE:FREQ=WEEKLY;UNTIL={until};BYDAY={day}"]

    return event

def main():
    print("Welcome to HKU Course Planner!")
    excel_reader = ExcelReader(EXCEL_FILE)
    calendar_manager = CalendarManager()
    directory_manager = DirectoryManager()

    while True:
        complete_course_list = excel_reader.read_excel()
        degree = list_menu_selector('degree', 'Select your degree:', list(degree_dict.keys()) + ['Exit'])

        if degree == 'Exit':
            break

        course_list = complete_course_list[complete_course_list['ACAD_CAREER'].str.contains(degree_dict[degree])]
        semester = list_menu_selector('semester', 'Select the semester:', ['Sem 1', 'Sem 2', 'Go back'])

        if semester == 'Go back':
            continue
        course_list = course_list[course_list['TERM'].str.contains(semester)]
        search_mode = list_menu_selector('search_mode', 'Select the search field:', list(search_mode_dict.keys()) + ['Go back'])

        if search_mode == 'Go back':
            continue

        while True:
            search = input(f"Enter the {search_mode.lower()} (-1 to go back): ")
            if search == '-1':
                break
            if search_mode == "Course code" and not re.match("[a-zA-Z]{4}[0-9]{4}", search):
                print("Invalid course code format. Please try again.")
                continue

            search_result = course_list[course_list[search_mode_dict[search_mode]].str.contains(search.upper())]
            if search_result.empty:
                print("No course found.")
                continue

            print(search_result)
            add_course = list_menu_selector('addcourse', 'Add this course to your Google Calendar?', ['Yes', 'No'])
            if add_course == 'Yes':
                mode = list_menu_selector('mode', 'Select mode:', ['One week', 'Whole semester'])
                add_to_calendar(calendar_manager, search_result, mode)

                make_dir = list_menu_selector('makedirectory', 'Create a folder for the added courses?', ['Yes', 'No'])
                if make_dir == 'Yes':
                    course = f"{search_result['COURSE CODE'].iloc[0]} - {search_result['COURSE TITLE OG'].iloc[0]}"
                    directory_manager.make_directory(semester, course)

            add_more = list_menu_selector('add_more', f"Search more courses by {search_mode.lower()}?", ['Yes', 'No'])
            if add_more == 'No':
                break

        if courses_added:
            clear = list_menu_selector('clear', 'Clear all entered events?', ['Yes', 'No'])
            if clear == 'Yes':
                for event_id, info in courses_added.items():
                    print(f"Deleting {info} from Google Calendar...")
                    calendar_manager.delete_event(event_id)
                courses_added.clear()

    print("Thank you for using HKU Course Planner!")

if __name__ == '__main__':
    main()