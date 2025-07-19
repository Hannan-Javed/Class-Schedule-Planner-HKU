import _pickle as pickle
import re, os
from excel_reader import ExcelReader
from calendar_manager import CalendarManager
from directory_manager import DirectoryManager
from config import EXCEL_FILENAME
from utils import list_menu_selector, select_courses, make_course_objects

degree_dict = {
    "UG - Undergraduate": 'UG',
    "TPG - Taught Postgraduate": 'TPG',
    "RPG - Research Postgraduate": 'RPG'
}
search_mode_dict = {
    "Course code": 'code',
    "Course title": 'title'
}

def main():
    print("Welcome to HKU Course Planner!")
    excel_reader = ExcelReader()
    if not os.path.exists(f'{EXCEL_FILENAME}.pkl'):
        
        complete_course_list = excel_reader.read_excel()
        
        courses = make_course_objects(complete_course_list)
        with open(f'{EXCEL_FILENAME}.pkl', 'wb') as f:
            pickle.dump(courses, f)
    else:
        with open(f'{EXCEL_FILENAME}.pkl', 'rb') as f:
            courses = pickle.load(f)
    
    calendar_manager = CalendarManager()
    directory_manager = DirectoryManager()

    while True:
        current_courses = []

        degree = list_menu_selector(
            'Select your degree:',
            list(degree_dict.keys()) + ['Exit']
        )

        if degree == 'Exit':
            break

        current_courses = [course for course in courses if course.degree == degree_dict[degree]]
        semester = list_menu_selector(
            'Select the semester:',
            ['Sem 1', 'Sem 2', 'Go back']
        )
        if semester == 'Go back':
            continue
        
        current_courses = [course for course in current_courses if semester in course.term]
        
        search_mode = list_menu_selector(
            'Select the search field:',
            list(search_mode_dict.keys()) + ['Go back']
        )
        if search_mode == 'Go back':
            continue

        while True:
            search = input(f"Enter the {search_mode.lower()} (-1 to go back): ")
            if search == '-1':
                break

            if search_mode == "Course code" and not re.match("[A-Za-z]{4}[0-9]{4}", search):
                print("Invalid course code format. Please try again.")
                continue

            search_result = [course for course in current_courses if search.upper() in getattr(course, search_mode_dict[search_mode]).upper()]

            if not search_result:
                print("No courses found. Please try again.")
                continue

            for i, course in enumerate(search_result):
                print(f"{i + 1}. {course.code} - {course.title}")

            add_course = list_menu_selector(
                'Do you want to add this course to your google calendar?' if len(search_result) == 1 else 'Do you want to add any of these courses to your google calendar?',
                ['Yes', 'No']
            )
            if add_course == 'Yes':
                search_result = select_courses(search_result) if len(search_result) > 1 else search_result

                add_course_title = (list_menu_selector(
                        'Do you want to add the course name in the event title?',
                        ['Yes', 'No']
                ) == "Yes") if locals().get('add_course_title') is None else add_course_title
                
                is_test_mode = list_menu_selector(
                    'Select mode:',
                    ['One week', 'Whole semester']
                ) == 'One week'
                
                for course in search_result:
                    if len(list(course.sections.keys())) > 0:
                        sections = course.select_sections()
                    else:
                        print("This course has no sections. Skipping")
                        continue
                    for section in sections:
                        if is_test_mode:
                            days_time_added = {}
                        for schedule_number, schedule in enumerate(course.sections[section]['schedules']):
                            if is_test_mode:
                                day = list(schedule.keys())[2]
                                if day in days_time_added.keys() and schedule[day] == days_time_added[day]:
                                    continue
                            calendar_manager.add_event(course.convert_to_calendar_event(section, add_course_title, is_test_mode, schedule_number))
                            if is_test_mode:
                                days_time_added[day] = schedule[day]

                make_dir = list_menu_selector(
                    'Create a folder for the course?' if len(search_result) == 1 else 'Create folders for the courses?',
                    ['Yes', 'No']
                ) if locals().get('make_dir') is None else make_dir
                if make_dir == 'Yes':
                    for course in search_result:
                        directory_manager.make_directory(course.term, course.code + ((' - ' + course.title) if add_course_title else '')) 

            add_more = list_menu_selector(
                f"Search more courses by {search_mode.lower()}?",
                ['Yes', 'No']
            )
            if add_more == 'No':
                break

    if calendar_manager.events_added:
        clear = list_menu_selector(
            'Clear all entered events?',
            ['Yes', 'No']
        )
        if clear == 'Yes':
            calendar_manager.clear_events()

    print("Thank you for using HKU Course Planner!")

if __name__ == '__main__':
    main()