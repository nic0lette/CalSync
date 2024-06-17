import argparse
import datetime
import os.path
import tomllib

from dateutil import parser as dt_parser
from google.auth.transport.requests import Request
from google.oauth2 import credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from ics import Event, Calendar

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['openid', 'https://www.googleapis.com/auth/calendar.readonly',
          'https://www.googleapis.com/auth/userinfo.email']


def google_event_to_ics_event(google_event) -> Event:
    event = Event()
    event.name = google_event['summary']
    event.begin = google_event['start'].get('dateTime', google_event['start'].get('date'))
    event.end = google_event['end'].get('dateTime', google_event['end'].get('date'))
    return event


def get_event_accepted(google_event, my_email: str) -> str:
    attendees = google_event.get("attendees", None)
    if attendees is None:
        return "unknown"
    for attendee in attendees:
        if attendee.get("email", "") == my_email:
            return attendee.get("responseStatus", "unknown")
    return "unknown"


def get_email_address(creds):
    userinfo_service = build('oauth2', 'v2', credentials=creds)
    user_info = userinfo_service.userinfo().get().execute()
    return user_info["email"]


def export_calendar(ical_file_path: str, config: dict):
    # Start with creds as None for PyCharm
    creds = None

    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = credentials.Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    user_email = get_email_address(creds)
    calendar_service = build('calendar', 'v3', credentials=creds)

    # Call the Calendar API
    gmt_timezone = datetime.timezone(datetime.timedelta(hours=0))
    now = datetime.datetime.now(gmt_timezone)
    time_min = (now - datetime.timedelta(days=config["export_range"]["days_ago"]))
    time_max = (now + datetime.timedelta(days=config["export_range"]["days"]))

    events_result = calendar_service.events().list(calendarId='primary',
                                                   timeMin=time_min.isoformat(),
                                                   timeMax=time_max.isoformat(),
                                                   singleEvents=True,
                                                   orderBy='startTime').execute()
    google_events = events_result.get('items', [])
    if len(google_events) == 0:
        return

    calendar = Calendar()
    for event in events_result.get('items', []):
        start = dt_parser.parse(event['start'].get('dateTime', event['start'].get('date')))
        end = dt_parser.parse(event['end'].get('dateTime', event['end'].get('date')))
        status = get_event_accepted(event, user_email)

        # Don't keep multi-day events
        if (end - start).days >= 1:
            continue

        # Filter out meals
        if event['summary'] in config["ignore_events"]:
            continue

        can_see_guests = event.get('guestsCanSeeOtherGuests', True)
        if not can_see_guests:
            large_event = True
        else:
            large_event = len(event.get('attendees', [])) >= config["large_event_size"]

        # A bit more processing for "large" events
        if large_event:
            # Only add large events that I've accepted or tentatively accepted
            if status != "accepted" and status != "tentative":
                continue

        ical_event = google_event_to_ics_event(event)
        calendar.events.add(ical_event)

    with open(ical_file_path, "w") as ical_file:
        ical_file.write(calendar.serialize())


def main():
    # Read config file
    with open("config.toml", "rb") as config_file:
        config = tomllib.load(config_file)

    # Parse arguments to the script
    parser = argparse.ArgumentParser(description="Calendar export script.")
    parser.add_argument('file_path', type=str, help='The path to the iCal file to create')
    args = parser.parse_args()
    export_calendar(args.file_path, config)


if __name__ == '__main__':
    main()
