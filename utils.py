import sys
import time
import threading
from PyInquirer import prompt
from functools import wraps

def list_menu_selector(field, prompt_message, choices):
    menu = prompt([
        {
            'type': 'list',
            'name': field,
            'message': prompt_message,
            'choices': choices,
        }
    ])
    return menu.get(field, None)

def loading_animation(message, stop_event):
    while not stop_event.is_set():
        for dots in range(4):  # 0 to 3 dots
            sys.stdout.write("\r" + message + " " * 4)
            sys.stdout.flush()
            sys.stdout.write("\r" + message + "." * dots)
            sys.stdout.flush()
            time.sleep(0.5)
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