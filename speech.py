from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from newsapi import NewsApiClient
import os
import time
import playsound
import speech_recognition as sr
from gtts import gTTS
import pytz
import subprocess
import requests
from bs4 import BeautifulSoup
from selenium import webdriver


SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
DAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
DAY_EXTENSIONS = ["nd", "rd", "th", "st"]

def speak(text):
    tts = gTTS(text=text, lang="en")
    filename = "voice.mp3"
    tts.save(filename)
    playsound.playsound(filename)
    os.remove(filename)

def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        said = ""

        try:
            said = r.recognize_google(audio)
            print(said)
        except Exception as e:
            return -1
            # print("Exception :" + str(e))

    return said

def authenticate_google():
    creds = None
    
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
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

            if int(start_time.split(":")[0]) < 12:  
                start_time = start_time + "am"
            else:
                start_time = str(int(start_time.split(":")[0])-12)  
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

def note(text):
    date  = datetime.datetime.now()
    file_name = str(date).replace(":","-") + "-note.txt"

    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe", file_name])


def open_something(text):
    text = text.lower()

    if "chrome" in text:
        speak("Opening google chrome")
        file_path = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

    elif "word" in text:
        speak("Opening Word")
        file_path = r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.exe" 

    elif "powerpoint" in text:
        speak("Opening Powerpoint")
        file_path = r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.exe"

    elif "excel" in text:
        speak("Opening Excel")
        file_path = r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.exe"

    elif "firefox" in text:
        speak("Opening Firefox")
        file_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

    elif "google" in text:
        speak("Opening google")
        browser = webdriver.Firefox()
        browser.get("http://www.google.com")
        return

    elif "facebook" in text:
        speak("Opening facebook")
        browser = webdriver.Firefox()
        browser.get("http://www.facebook.com")
        return

    elif "youtube" in text:
        speak("Opening Youtube")
        browser = webdriver.Firefox()
        browser.get("http://www.youtube.com")
        return

    elif "gmail" in text:
        speak("Opening Gmail")
        browser = webdriver.Firefox()
        browser.get("http://www.gmail.com")
        return
        
    subprocess.Popen(file_path)
    

def get_news():
    api_key = os.environ.get('NEWS_API_KEY')
    newsapi = NewsApiClient(api_key=api_key)
    
    top_headlines = newsapi.get_top_headlines(sources='bbc-news,the-verge,cnn',
                                            language='en')
                                            
    for i in range(min(len(top_headlines['articles']),5)):
        print(top_headlines['articles'][i]['description'])
        speak(top_headlines['articles'][i]['description'])


SERVICE = authenticate_google()
i = 0

while(True):
    i+= 1 
    if i==1:
        print("Try saying 'Hello Assistant' ")
    else:
        speak("Do you want me to do anyting else?")

    text = get_audio()
    if text==-1:
        break
    if "no" in text:
        speak("Okay")
        break

    text = text.lower()
    if "hello" or "assistant" or "hello assistant" in text:
        speak("How can I help you?")
        text = get_audio()

    if text==-1:
        break

    #Goole Calender
    CALENDAR_STRS = ["what do i have", "do i have", "am i busy"]
    for phrase in CALENDAR_STRS:
        if phrase in text.lower():
            date = get_date(text)
            if date:
                get_events(date, SERVICE)
            else:
                speak("Sorry,I don't understand")


    #Making a note in notepad
    NOTE_STRS = ["make a note", "write this down", "remember this", "type this"]
    for phrase in NOTE_STRS:
        if phrase in text.lower():
            speak("What would you like me to write down? ")
            write_down = get_audio()
            note(write_down)
            speak("I have made a note of that.")

    #Opening something
    if "open" in text.lower():
        open_something(text)


    #Showing news
    if "news" in text.lower():
        get_news()




