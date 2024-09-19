import os
import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

class GoogleCalendar:
    SCOPES = ['https://www.googleapis.com/auth/calendar']

    def __init__(self):
        self.creds = self.get_credentials()
        self.service = build("calendar", "v3", credentials=self.creds)

    def get_credentials(self):
        creds = None
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", self.SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json", self.SCOPES)
                creds = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(creds.to_json())
        return creds

    def list_events(self, time_min=None, time_max=None, max_results=10):
        if time_min is None:
            time_min = datetime.datetime.utcnow().isoformat() + 'Z'
        events_result = self.service.events().list(calendarId='primary', timeMin=time_min, 
                                                   timeMax=time_max, maxResults=max_results, 
                                                   singleEvents=True, orderBy='startTime').execute()
        return events_result.get('items', [])

    def create_event(self):
        event_details = self.collect_event_details()
        event = self.format_event(event_details)
        created_event = self.service.events().insert(calendarId='primary', body=event).execute()
        print(f'Event created: {created_event.get("htmlLink")}')

    def collect_event_details(self):
        summary = input("Enter the event summary: ")
        description = input("Enter the event description: ")
        location = input("Enter the event location: ")
        start_time = self.collect_datetime("start")
        end_time = self.collect_datetime("end")
        return {'summary': summary, 'description': description, 'location': location, 
                'start_time': start_time, 'end_time': end_time}

    def collect_datetime(self, type):
        while True:
            date_str = input(f"Enter the {type} date (DD-MM-YYYY): ")
            try:
                hour = int(input(f"Enter the {type} hour (0-23): "))
                minute = int(input(f"Enter the {type} minute (0-59): "))
                date = datetime.datetime.strptime(date_str, "%d-%m-%Y")
                return date.replace(hour=hour, minute=minute).isoformat() + 'Z'
            except ValueError:
                print("Invalid date format. Please enter the date as DD-MM-YYYY.")
            except Exception as e:
                print(f"Error: {e}. Please try again.")

    def format_event(self, details):
        return {
            'summary': details['summary'],
            'location': details['location'],
            'description': details['description'],
            'start': {'dateTime': details['start_time'], 'timeZone': 'UTC'},
            'end': {'dateTime': details['end_time'], 'timeZone': 'UTC'},
        }

    def update_event(self):
        # Allowing to update both past and future events
        events = self.list_events(time_max=datetime.datetime.utcnow().isoformat() + 'Z', max_results=20)
        if not events:
            print("No events found.")
            return
        self.display_events(events)
        event_index = int(input("Enter the event number you want to update: ")) - 1
        if 0 <= event_index < len(events):
            event_id = events[event_index]['id']
            event_to_update = self.service.events().get(calendarId='primary', eventId=event_id).execute()
            updated_details = self.collect_event_details()
            updated_event = self.format_event(updated_details)
            self.service.events().update(calendarId='primary', eventId=event_id, body=updated_event).execute()
            print("Event updated successfully.")
        else:
            print("Invalid event number.")

    def display_events(self, events):
        print("Events:")
        for i, event in enumerate(events):
            start = event["start"].get("dateTime", event["start"].get("date"))
            print(f"{i + 1}: {start} - {event['summary']} (ID: {event['id']})")

    def delete_event(self):
        events = self.list_events(time_max=datetime.datetime.utcnow().isoformat() + 'Z', max_results=20)
        if not events:
            print("No events found.")
            return
        self.display_events(events)
        event_index = int(input("Enter the event number you want to delete: ")) - 1
        if 0 <= event_index < len(events):
            event_id = events[event_index]['id']
            self.service.events().delete(calendarId='primary', eventId=event_id).execute()
            print("Event deleted successfully.")
        else:
            print("Invalid event number.")

    def retrieve_events_in_month(self):
        year = int(input("Enter the year (YYYY): "))
        month = int(input("Enter the month (1-12): "))
        time_min, time_max = self.get_month_dates(year, month)
        events = self.get_month_events(time_min, time_max)
        if not events:
            print("No upcoming events found for this month.")
            return
        self.display_month_events(events, year, month)

    def get_month_dates(self, year, month):
        start_date = datetime.datetime(year, month, 1)
        end_date = datetime.datetime(year + (month // 12), (month % 12) + 1, 1)
        return start_date.isoformat() + 'Z', end_date.isoformat() + 'Z'

    def get_month_events(self, time_min, time_max):
        events = []
        page_token = None
        while True:
            events_result = self.service.events().list(calendarId='primary', timeMin=time_min,
                                                       timeMax=time_max, pageToken=page_token,
                                                       singleEvents=True, orderBy='startTime').execute()
            events.extend(events_result.get('items', []))
            page_token = events_result.get('nextPageToken')
            if not page_token:
                break
        return events

    def display_month_events(self, events, year, month):
        print(f"Upcoming events for {month}/{year}:")
        for event in events:
            start = event["start"].get("dateTime", event["start"].get("date"))
            print(start, event["summary"])

    def input_function(self):
        while True:
            print("\nGoogle Calendar Operations:")
            print("1. Create a new event")
            print("2. Update an existing event")
            print("3. Delete an event")
            print("4. Retrieve upcoming events in a specific month")
            print("5. Exit")

            choice = input("Enter your choice (1-5): ")
            if choice == '1':
                self.create_event()
            elif choice == '2':
                self.update_event()
            elif choice == '3':
                self.delete_event()
            elif choice == '4':
                self.retrieve_events_in_month()
            elif choice == '5':
                print("Exiting the program.")
                break
            else:
                print("Invalid choice. Please try again.")

def main():
    calendar = GoogleCalendar()
    calendar.input_function()

if __name__ == "__main__":
    main()
