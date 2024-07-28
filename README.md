# HKU Course Planner
This Python script helps HKU students add their courses to their Google Calendar. It reads course information from an Excel file and allows users to search for courses by code or title, then add them to their calendar.
## Features
- **Course Import**: Reads course information from an Excel file (currently 2024-25_class_timetable_20240722.xlsx, can be changed in global variables).
- **Degree Selection**: Allows users to choose their degree program (UG, TPG, RPG).
- **Semester Selection**: Allows users to choose the semester (Sem 1 or Sem 2).
- **Course Search**: Searches for courses by code or title.
- **Calendar Integration**: Adds selected courses to the user's Google Calendar.
- **Recurrence Options**: Allows users to choose between adding courses for one week (e.g. for seeing clashes) or the whole semester.
## Installation
1. Install the required packages:<br>
`pip install -r requirements.txt`
2. Enable the Google Calendar API for your project.
3. Create API credentials (OAuth 2.0 client ID) for your project.
4. Download the credentials.json file and place it in the same directory as the Python script.
5. Run the script. It will guide you through the authentication process and create a token.json file to store your authentication credentials.
## Usage
Run `main.py` to launch the program.

