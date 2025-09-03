from gmail_api import init_gmail_service, get_email_messages, get_email_message_details, download_attachments
from PyPDF2 import PdfReader
import os
from googleapiclient.errors import HttpError
from datetime import datetime, timezone, timedelta
from calendar_api import init_calendar_service
import time
import re

#!/usr/bin/python3

client_file = "client_secret.json"
try:
    mail_service = init_gmail_service(client_file)
    calendar_service = init_calendar_service(client_file)

except TokenExpiredException:
    print('Tokens have expired: error 401')
    os.remove('ENTER_TOKEN_FILE_PATH_HERE')
    mail_service = init_gmail_service(client_file)
    calendar_service = init_calendar_service(client_file)

except Exception as e:
    print(f'An unexpected error has occured: {e}')
messages = get_email_messages(mail_service, max_results=100)

path = 'ENTER_DOWNLOAD_DIRECTORY_PATH_HERE'

count = 0
for msg in messages:
    count += 1
    print(f'Emails read: {count}', end='\r', flush=True)
    detal = detail = get_email_message_details(mail_service, msg['id'])
    if "ENTER_SENDER_EMAIL_HERE" in detail['sender']:
        print('Email found')
        download_attachments(mail_service, 'me', msg['id'], path)
        break

dir_list = os.listdir(path)
print(dir_list)
pdf_list = []
amount = 0

for item in dir_list:
    if item.endswith('.pdf'):
        pdf_list.append(item)
        amount += 1

for i in range(amount):
    text = []
    with open(f'{path}/{pdf_list[i]}', 'rb') as f:
        pdf_reader = PdfReader(f)
        num_page = len(pdf_reader.pages)

        for page_num in range(num_page):
            page = pdf_reader.pages[page_num]
            page_text = page.extract_text()
            text.extend(page_text.split('\n'))
    
    combined_time_data = []
    cleaned_dates = []

    for row in text:
        if isinstance(row, str):
            row = row.split(' ')

        if 'ENTER_YOUR_NAME_IDENTIFIER_HERE' in row[0]:
            time_data = row[:-1]
            j = 0
            while j < len(time_data):
                if time_data[j] in ['RO', 'OFF']:
                    combined_time_data.append(time_data[j])
                    combined_time_data.append('-')
                elif time_data[j] == '0.0':
                    if len(combined_time_data) == 0 or combined_time_data[-1] != '-':
                        combined_time_data.append('-')
                        combined_time_data.append('-')
                    combined_time_data.append(time_data[j])
                elif time_data[j] in ['AM', 'PM'] and len(combined_time_data) > 0:
                    combined_time_data[-1] = f"{combined_time_data[-1]} {time_data[j]}"
                else:
                    combined_time_data.append(time_data[j])
                j += 1

        for item in row:
            match = re.match(r'\d{1,2}/\d{1,2}/\d{4}', item)
            if match:
                cleaned_dates.append(match.group())
        
    for i in range(len(cleaned_dates)):
        temp = cleaned_dates[i].split('/')
        y = temp[2]
        m = temp[0]
        if len(m) < 2:
            m = '0' + m
        d = temp[1]
        if len(d) < 2:
            d = '0' + d

        day = y + '-' + m + '-' + d
        cleaned_dates[i] = day

    for i in range(len(combined_time_data) - 1, -1, -1):
        if 'PM' in combined_time_data[i]:
            time_t = combined_time_data[i].replace('PM', '').strip()
            time_t = time_t.split(':')
            hour = int(time_t[0]) + 12
            if hour == 24:
                hour = 12
            minute = int(time_t[1])
            combined_time_data[i] = f"{hour}:{minute:02d}"

        elif 'AM' in combined_time_data[i]:
            time_t = combined_time_data[i].replace('AM', '').strip()
            time_t = time_t.split(':')
            hour = int(time_t[0])
            minute = int(time_t[1])
            combined_time_data[i] = f"{hour}:{minute:02d}"

        if i % 3 == 0:
            combined_time_data.pop(i)

    time_groups = []

    for i in range(0, len(combined_time_data), 2):
        if i + 1 < len(combined_time_data):
            time_groups.append([combined_time_data[i], combined_time_data[i + 1]])
        else:
            time_groups.append([combined_time_data[i]])

    print(time_groups)

    offset_seconds = -time.timezone if time.localtime().tm_isdst == 0 else -time.altzone
    local_offset = timezone(timedelta(seconds=offset_seconds))

    for i in range(len(time_groups)):
        if time_groups[i][0] in ['RO', 'OFF', '-']:
            print(f"Skipping invalid time group: {time_groups[i]}")
            continue
        if len(time_groups[i]) < 2:
            print(f"Skipping incomplete time group: {time_groups[i]}")
            continue
        if len(cleaned_dates) <= i:
            print(f"Skipping time group due to missing date: {time_groups[i]}")
            continue

        start_dt = datetime.fromisoformat(f"{cleaned_dates[i]}T{time_groups[i][0]}:00").replace(tzinfo=local_offset)
        end_dt = datetime.fromisoformat(f"{cleaned_dates[i]}T{time_groups[i][1]}:00").replace(tzinfo=local_offset)

        if start_dt >= end_dt:
            print(f"Skipping invalid time range: start_time={start_dt}, end_time={end_dt}")
            continue

        start_time = start_dt.isoformat()
        end_time = end_dt.isoformat()

        try:
            events_result = calendar_service.events().list(
                calendarId='primary',
                timeMin=start_time,
                timeMax=end_time,
                singleEvents=True,
                orderBy='startTime',
                q='Work'
            ).execute()

            duplicate_found = False
            for existing_event in events_result.get('items', []):
                existing_start = datetime.fromisoformat(existing_event['start']['dateTime'])
                existing_end = datetime.fromisoformat(existing_event['end']['dateTime'])

                if existing_event.get('summary', '').strip().lower() == 'work' and \
                existing_start == start_dt and existing_end == end_dt:
                    print(f"Skipping duplicate event: {start_dt} to {end_dt}")
                    duplicate_found = True
                    break

            if duplicate_found:
                continue

        except HttpError as error:
            print(f"An error occurred while checking for duplicates: {error}")
            continue

        event = {
            'summary': 'Work',
            'start': {
                'dateTime': start_time,
                'timeZone': 'America/Chicago'
            },
            'end': {
                'dateTime': end_time,
                'timeZone': 'America/Chicago'
            },
        }

        try:
            print(f"Creating event: {event}")
            event_result = calendar_service.events().insert(calendarId='primary', body=event).execute()
            print(f"Event created: {event_result.get('htmlLink')}")
        except HttpError as error:
            print(f"An error occurred: {error}")

for i in range(len(pdf_list) - 1, -1, -1):
    file_path = f'{path}/{pdf_list[i]}'
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f'Removed {file_path}')

print(r'''
                                                 
                                                 
`7MM"""Yb.     .g8""8q. `7MN.   `7MF'`7MM"""YMM  
  MM    `Yb. .dP'    `YM. MMN.    M    MM    `7  
  MM     `Mb dM'      `MM M YMb   M    MM   d    
  MM      MM MM        MM M  `MN. M    MMmmMM    
  MM     ,MP MM.      ,MP M   `MM.M    MM   Y  , 
  MM    ,dP' `Mb.    ,dP' M     YMM    MM     ,M 
.JMMmmmdP'     `"bmmd"' .JML.    YM  .JMMmmmmMMM 
                                                 
                                                 
''')
