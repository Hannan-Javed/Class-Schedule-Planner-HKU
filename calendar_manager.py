import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/calendar"]

class CalendarManager:
    def __init__(self):
        self.service = self.get_credentials()
        self.events_added = []

    def get_credentials(self):
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        try:
            service = build('calendar', 'v3', credentials=creds)
            return service
        except HttpError as e:
            print("An error occurred: ", e)
            return None

    def add_event(self, event):
        try:
            event = self.service.events().insert(calendarId='primary', body=event).execute()
            self.events_added.append(event['id'])
            return event
        except HttpError as e:
            print("Failed to add event:", e)
            return None
    
    def delete_event(self, event_id):
        try:
            self.service.events().delete(calendarId='primary', eventId=event_id).execute()
            self.events_added.remove(event_id)
        except HttpError as e:
            print("Failed to delete event:", e)

    def clear_events(self):
        for event_id in self.events_added:
            self.delete_event(event_id)
        self.events_added = []