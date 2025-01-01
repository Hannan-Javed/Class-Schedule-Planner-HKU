from datetime import timedelta
import pandas as pd

day_index = {'MON': 0, 'TUE': 1, 'WED': 2, 'THU': 3, 'FRI': 4}

class Course:
    def __init__(self, row):
        self.start_date = row['START DATE'].date()
        self.end_date = row['END DATE'].date()
        self.code = row['COURSE CODE']
        self.title = row['COURSE TITLE OG']
        self.section = row['CLASS SECTION']
        self.description = row['VENUE']
        self.day = [day[:2] for day in day_index if pd.notnull(row[day])][0]
        self.days_difference = day_index[[day for day in day_index if pd.notnull(row[day])][0]]
        self.term = row['TERM']
        self.start_time = row['START TIME']
        self.end_time = row['END TIME']
    
    def create_event(self, testmode, add_section_title):
        start_date = self.start_date
        if testmode == 'One week':
            today = pd.Timestamp('today').date()
            if 'Sem 1' in self.term:
                # add to next week
                start_date = today + timedelta(days=(7 - today.weekday()))
            elif 'Sem 2' in self.term:
                # add to next next week
                start_date = today + timedelta(days=(14 - today.weekday()))
            start_date += timedelta(days=self.days_difference)
        
        event = {
            'summary': f"{self.code} {self.title} {(' ' + self.section) if add_section_title == 'Yes' else ''}",
            'description': self.description,
            'start': {
                'dateTime': f"{start_date}T{self.start_time}+08:00",
                'timeZone': 'Asia/Hong_Kong',
            },
            'end': {
                'dateTime': f"{start_date}T{self.end_time}+08:00",
                'timeZone': 'Asia/Hong_Kong',
            },
            'colorId': 7,
            'reminders': {
                'useDefault': False,
                'overrides': [],
            },
        }
        
        if testmode == 'Whole semester':
            until = int(str(self.end_date).replace('-', '')) + 1
            event['recurrence'] = [f"RRULE:FREQ=WEEKLY;UNTIL={until};BYDAY={self.day[:2]}"]
        
        return event