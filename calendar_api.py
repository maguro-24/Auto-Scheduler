from google_api import create_service
from datetime import datetime

def init_calendar_service(client_file, api_name='calendar', api_version='v3', scopes=['https://www.googleapis.com/auth/calendar']):
    return create_service(client_file, api_name, api_version, scopes)

def add_event(service, calendar_id='primary', dates=[], times=[], summary='No Title', location='', description='', timezone='America/New_York'):
    created_events = []

    for date_str, time_pair in zip(dates, times):
        start_time_str, end_time_str = time_pair

        # Skip invalid entries
        if start_time_str in ['-', 'RO'] or end_time_str in ['-', 'RO']:
            continue

        try:
            # Combine date and time strings into datetime objects
            start_dt = datetime.strptime(f"{date_str} {start_time_str}", "%Y-%m-%d %H:%M")
            end_dt = datetime.strptime(f"{date_str} {end_time_str}", "%Y-%m-%d %H:%M")

            event = {
                'summary': summary,
                'start': {
                    'dateTime': start_dt.isoformat(),
                    'timeZone': timezone,
                },
                'end': {
                    'dateTime': end_dt.isoformat(),
                    'timeZone': timezone,
                },
            }

            response = service.events().insert(calendarId=calendar_id, body=event).execute()
            created_events.append(response.get('htmlLink'))

        except Exception as e:
            print(f"Failed to create event on {date_str} {start_time_str}–{end_time_str}: {e}")

    return created_events