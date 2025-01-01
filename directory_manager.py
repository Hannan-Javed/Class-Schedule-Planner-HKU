import os
from tkinter import filedialog
from utils import list_menu_selector  # Added import

class DirectoryManager:
    def get_default_download_directory(self):
        home_directory = os.path.expanduser("~")
        return os.path.join(home_directory, 'Downloads') if os.name in ['posix', 'nt'] else None

    def make_directory(self, term, course):
        choice = list_menu_selector('directory', 'Please select the directory to save the files:', ['Downloads', 'Input Directory Path'])
        directory = self.get_default_download_directory() if choice == 'Downloads' else self.prompt_directory()
        if directory:
            directory = os.path.join(directory, term, course)
            os.makedirs(directory, exist_ok=True)
            return directory
        return None

    def prompt_directory(self):
        while True:
            directory = input("Please enter the directory path: ")
            if os.path.exists(directory):
                return directory
            print("The directory does not exist. Please try again.")