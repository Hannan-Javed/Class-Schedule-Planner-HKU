import sys, time, threading
from PyInquirer import prompt
from functools import wraps
from course import Course

def list_menu_selector(prompt_message, choices):
    menu = prompt([
        {
            'type': 'list',
            'name': 'choice',
            'message': prompt_message,
            'choices': choices,
        }
    ])
    return menu['choice']

def loading_animation(message, stop_event):
    spinner = ['|', '/', '-', '\\']
    idx = 0
    while not stop_event.is_set():
        sys.stdout.write("\r" + message + " " + spinner[idx])
        sys.stdout.flush()
        idx = (idx + 1) % len(spinner)
        time.sleep(0.1)
    print()

def with_loading_animation(message):
    """
    Decorator to display a loading animation while a function is executing.
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            stop_event = threading.Event()
            animation_thread = threading.Thread(target=loading_animation, args=(message, stop_event), daemon=True)
            animation_thread.start()
            try:
                result = func(*args, **kwargs)
            finally:
                stop_event.set()
                animation_thread.join()
            return result
        return wrapper
    return decorator

@with_loading_animation("Making courses list")
def make_course_objects(course_list):
    courses = []
    grouped = course_list.groupby(["TERM", "COURSE CODE"])
    for _, group in grouped:
        course_obj = Course(group)
        courses.append(course_obj)

    return courses

def select_courses(courses):
    courses_to_add = input(f"Enter courses to add (comma-separated): ").split(',')
    while any(int(course_num)<1 or int(course_num)>len(courses) for course_num in courses_to_add):
        courses_to_add = input(f"Invalid courses. Please enter valid courses: ").split(',')
    s = []
    for course_num in courses_to_add:
        s.append(courses[int(course_num)-1])

    return s
        
