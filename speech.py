from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import time
import playsound
import speech_recognition as sr
from gtts import gTTS
import pytz

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
DAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
DAY_EXTENSIONS = ["nd", "rd", "th", "st"]

def speak(text):
    tts = gTTS(text=text, lang="en")
    filename = "voice.mp3"
    tts.save(filename)
    playsound.playsound(filename)

def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        said = ""

        try:
            said = r.recognize_google(audio)
        except Exception as e:
            print("Exception :" + str(e))

    return said

def authenticate_google():
    creds = None
    
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    return service


def get_events(day, service):
    
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)

    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(),
                                        timeMax=end_date.isoformat(), singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No events found.')
    else:
        if(len(events)==1):
            speak(f"You have {len(events)} event.")
        else:
            speak(f"You have {len(events)} events.")

        for i,event in enumerate(events):
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])
            start_time = str(start.split("T")[1].split("-")[0])[:2]

            if int(start_time.split(":")[0]) < 12:  # if the event is in the morning
                start_time = start_time + "am"
            else:
                start_time = str(int(start_time.split(":")[0])-12)  # convert 24 hour time to regular
                start_time = start_time + "pm"  

            print(start_time)
            tts = gTTS(event["summary"] + " at " + start_time)
            file = str("file" + str(i) + ".mp3")
            tts.save(file)
            playsound.playsound(file)
            

def get_date(text):
    text = text.lower()
    today = datetime.date.today()

    if text.count("today") > 0:
        return today

    if text.count("tomorrow") > 0:
        return datetime.date(month= today.month, day=today.day+1, year=today.year)

    day = -1
    day_of_week = -1
    month = -1
    year = today.year
    
    for word in text.split():
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        elif word in DAYS:
            day_of_week = DAYS.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for ext in DAY_EXTENSIONS:
                found = word.find(ext)
                if found > 0:
                    try:
                        day = int(word[:found])
                    except:
                        pass
         
    if month < today.month and month != -1:
        year = year + 1

    if day < today.day and month == -1 and day != -1:
        month = month + 1

    if month == -1 and day == -1 and day_of_week != -1:
        current_day_of_week = today.weekday()
        diff = day_of_week - current_day_of_week

        if diff < 0:
            diff += 7
            if text.count("next") > 0:
                diff += 7

        return today + datetime.timedelta(diff)

    if month == -1 or day == -1:
        return None

    return datetime.date(month=month, day=day, year=year)

print("Say OK Assistant")
first = get_audio()
print(first)
if "ok assistant" or "assistant" or "ok" in first.lower():
    speak("Hey there! How can I help you?")

SERVICE = authenticate_google()
text = get_audio()

CALENDAR_STRS = ["what do i have", "do i have plans", "am i busy"]
for phrase in CALENDAR_STRS:
    if phrase in text.lower():
        date = get_date(text)
        if date:
            get_events(date, SERVICE)
        else:
            speak("Please Try Again")
