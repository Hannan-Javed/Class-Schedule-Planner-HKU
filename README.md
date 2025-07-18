# HKU Course Planner
This Python script helps HKU students add their courses to their Google Calendar. It reads course information from an Excel file and allows users to search for courses by code or title, then add them to their calendar and create directories for that course.
## Features
- **Course Import**: Reads course information from an Excel file.
- **Degree Selection**: Allows users to choose their degree program (UG, TPG, RPG).
- **Semester Selection**: Allows users to choose the semester.
- **Course Search**: Searches for courses by code or title.
- **Calendar Integration**: Adds selected courses to the user's Google Calendar.
- **Recurrence Options**: Allows users to choose between adding courses for one week for testing (e.g. for seeing clashes) or the whole semester.
- **Deletion**: Allow users to delete all entered courses in their google calendar.
- **Make Course Directories**: Allow users to create directories for courses.
- **Optimized Loading Times**: Saves objects (courses) using pickle for faster load times during successice execution(s).
## Installation
1. Clone the repository
    ```
    git clone https://github.com/Hannan-Javed/Class-Schedule-Planner-HKU
    ```
2. Install the required packages:
    ```
    pip install -r requirements.txt
    ```
3. Create a Google Cloud Console Project and enable the Google Calendar API for your project.
4. Create API credentials (OAuth 2.0 client ID) for your project.
5. Download the credentials.json file in the same directory as the Python script.
6. Download the class schedule excel file through HKU Portal > SIS > Timetables > Class Timetable also in the same directory.

You can also refer to the guideline here for Google project setup:<br>
https://www.youtube.com/watch?v=B2E82UPUnOY&t=463s
## Usage
- Change excel filename in `config.py` to the one downloaded
- Run `main.py`