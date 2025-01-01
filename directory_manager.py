import os
from tkinter import filedialog
from utils import list_menu_selector  # Added import

class DirectoryManager:
    def __init__(self):
        self.directory = None
    def get_default_download_directory(self):
        home_directory = os.path.expanduser("~")
        return os.path.join(home_directory, 'Downloads') if os.name in ['posix', 'nt'] else None

    def make_directory(self, term, course):
        if not self.directory:
            choice = list_menu_selector('directory', 'Please select the directory to save the files:', ['Downloads', 'Input Directory Path'])
            self.directory = self.get_default_download_directory() if choice == 'Downloads' else self.prompt_directory()
        dir = os.path.join(self.directory, term, course)
        os.makedirs(dir, exist_ok=True)
        

    def prompt_directory(self):
        while True:
            directory = input("Please enter the directory path: ")
            if os.path.exists(directory):
                return directory
            print("The directory does not exist. Please try again.")