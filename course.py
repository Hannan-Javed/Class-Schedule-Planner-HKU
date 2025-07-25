from datetime import timedelta
import pandas as pd

day_index = {'MON': 0, 'TUE': 1, 'WED': 2, 'THU': 3, 'FRI': 4}

class Course:

    def __init__(self, group_df):
        self.code = group_df.iloc[0]['COURSE CODE']
        self.title = group_df.iloc[0]['COURSE TITLE']
        self.term = group_df.iloc[0]['TERM']
        self.degree = group_df.iloc[0]['ACAD_CAREER']
        
        self.sections = {}
        grouped = group_df.groupby(['CLASS SECTION'])
        for section, sub_df in grouped:
            schedule_list = []
            for _, row in sub_df.iterrows():
                schedule = {
                    'venue': row['VENUE'],
                    'start_date': row['START DATE'].date(),
                    'end_date': row['END DATE'].date()
                }
                # Each row should only have one valid day column.
                day_found = next((day for day in day_index if pd.notnull(row[day])), None)
                if day_found:
                    schedule[day_found] = {
                        'start_time': row['START TIME'],
                        'end_time': row['END TIME']
                    }
                
                schedule_list.append(schedule)
            
            self.sections[section[0]] = schedule_list

    def select_sections(self):
        sections = list(self.sections.keys())
        if len(sections) == 1:
            return sections
        sections_to_add = input(
            f"There are multiple sections for {self.code}: {self.title}. "
            f"Enter sections to add (comma-separated): {', '.join(sections)}\nSections: "
        ).upper()
        while not all(section in sections for section in sections_to_add.split(',')):
            sections_to_add = input(
                f"Invalid sections. Please enter valid sections: {', '.join(sections)}\nSections: "
            ).upper()
        
        return [section for section in self.sections if section in sections_to_add.split(',')]
    
    def convert_to_calendar_event(self,section_name, add_section_title, testmode, schedule_num):
        
        if section_name is None:
            section_name = next(iter(self.sections))

        section = self.sections[section_name]
        
        schedule = section[schedule_num]
        start_date = schedule['start_date']
        
        day, time = list(schedule.items())[3]
        
        # Adjust start_date if in "One week" (test) mode.
        if testmode:
            today = pd.Timestamp('today').date()
            start_date = today + timedelta(days=(7 + (7 if 'Sem 2' in self.term else 0) - today.weekday()))
            start_date += timedelta(days=day_index.get(day, 0))
        
        summary = f"{self.code}{(' - ' + section_name) if len(self.sections) > 1 else ''}{(' - ' + self.title) if add_section_title else ''}"
        
        event = {
            'summary': summary,
            'description': schedule['venue'],
            'start': {
                'dateTime': f"{start_date}T{time['start_time']}+08:00",
                'timeZone': 'Asia/Hong_Kong'
            },
            'end': {
                'dateTime': f"{start_date}T{time['end_time']}+08:00",
                'timeZone': 'Asia/Hong_Kong'
            },
            'colorId': 7,
            'reminders': {
                'useDefault': False,
                'overrides': []
            }
        }
        
        # For whole semester mode, add recurrence based on end_date.
        if not testmode:
            until = int(str(schedule['end_date']).replace('-', '')) + 1
            event['recurrence'] = [f"RRULE:FREQ=WEEKLY;UNTIL={until};BYDAY={day[:2]}"]
        return event


# Structure of the course data
# Build a dictionary of sections
# sections = {
#   section1: [
#               {
#                   'venue': <venue>,
#                   'start_date': s1, 
#                   'end_date': e1,
#                   'MON': {'start_time': t1, 'end_time': t1_end},
#               },
#               {
#                   'venue': <venue>,
#                   'start_date': s2,
#                   'end_date': e2,
#                   'TUE': {'start_time': t2, 'end_time': t2_end},
#               },
#               ...
#             ],
#   },
#   section2: { ... }
# }