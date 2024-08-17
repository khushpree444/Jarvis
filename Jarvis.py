import ctypes
from email import generator
import json
import subprocess
from bs4 import BeautifulSoup
import pyautogui
import pyttsx3
import requests
import speech_recognition as sr
import datetime
import time
import os
import random
import wikipedia
import webbrowser
import pywhatkit
import pyjokes
import cv2
import psutil
import requests
import pygame


engine = pyttsx3.init('sapi5')

voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in') # type: ignore
        print(f'User said: {query}\n')

    except Exception as e:
        print("Say that again please...")
        return "None"
    return query

def wish():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning sir")
    elif hour>=12 and hour<18:
        speak("Good Afternoon sir")
    else:
        speak("Good Evening sir")
    speak("I am Jarvis. How may I help you?")

if __name__ == '__main__':
    wish()
    while True:
        query = takeCommand().lower()
        if "Jarvis" in query:
            query = query.replace("Jarvis", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia.........")
            print(results)

        if 'open notepad' in query or "opening notepad" in query:
            speak("opening notepad")
            npath = "C:\\Windows\\System32\\notepad.exe"
            os.startfile(npath)
        
        elif 'close notepad' in query or "closing notepad" in query:
            speak("okay sir, closing notepad")
            os.system("taskkill /f /im notepad.exe")

        elif 'open command prompt' in query or 'open cmd' in query:
            speak("okay sir, opening command prompt")
            path = "C:\\Windows\\System32\\cmd.exe"
            os.system("start cmd")

        elif 'close command prompt' in query or 'close cmd' in query:
            speak("okay sir, closing command prompt")
            os.system("taskkill /f /im cmd.exe")

        elif 'open code' in query or 'open visual studio code' in query or 'open vs code' in query:
            speak("okay sir, opening code")
            path = "C:\\Users\\khush\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(path)

        elif 'close code' in query or 'close visual studio code' in query or 'close vs code' in query:
            speak("okay sir, closing code")
            os.system("taskkill /f /im Code.exe")


        elif 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)
        
        elif 'open youtube' in query or 'open yt' in query or "opening youtube" in query:
            webbrowser.open("youtube.com")

        elif"close youtube" in query or "closing youtube" in query:
            speak("okay sir, closing youtube")
            os.system("taskkill /f /im msedge.exe")

        elif 'open instagram' in query or 'open insta' in query or "opening instagram" in query:
            webbrowser.open("instagram.com")

        elif"close instagram" in query or "closing instagram" in query:
            speak("okay sir, closing instagram")
            os.system("taskkill /f /im msedge.exe")

        elif 'open facebook' in query or 'open fb' in query or "opening facebook" in query:
            webbrowser.open("www.facebook.com")

        elif"close facebook" in query or "closing facebook" in query:
            speak("okay sir, closing facebook")
            os.system("taskkill /f /im msedge.exe")

        elif 'open google' in query  or "opening google" in query:
            speak("sir, what should i search on google?")
            cm = takeCommand().lower()
            webbrowser.open(f"https://www.google.com/search?q={cm}")
            webbrowser.open(f"{cm}")

        elif"close google" in query or "closing google" in query:
            speak("okay sir, closing google")
            os.system("taskkill /f /im msedge.exe")

        elif "search on google" in query:
            speak("What do you want to search on google?")
            cm = takeCommand().lower()
            webbrowser.open(f"https://www.google.com/search?q={cm}")
            webbrowser.open(f"{cm}")

        elif 'search on youtube' in query:
            speak("What do you want to search on youtube?")
            cm = takeCommand().lower()
            webbrowser.open(f"https://www.youtube.com/results?search_query={cm}")
            webbrowser.open(f"{cm}")

        elif"close youtube" in query or "closing youtube" in query:
            speak("okay sir, closing youtube")
            os.system("taskkill /f /im msedge.exe")

        elif "play on youtube" in query:
            speak("What do you want to play on youtube?")
            cm = takeCommand().lower()
            pywhatkit.playonyt(f"{cm}") # type: ignore
            speak("playing on youtube")

        elif 'the time' in query or 'time' in query or 'time please' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif "the day" in query or  "today" in query or "day" in query:
            strDay = datetime.datetime.now().strftime("%A")
            speak(f"Sir, today is {strDay}")

        elif 'open word' in query or 'open word document' in query or "word" in query:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            os.startfile(path)

        elif 'close word' in query or 'close word document' in query or "close word" in query:
            speak("okay sir, closing word")
            os.system("taskkill /f /im WINWORD.EXE")

        elif 'open excel' in query or 'open excel sheet' in query or "opeining excel" in query:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            os.startfile(path)

        elif 'close excel' in query or 'close excel sheet' in query or "closeing excel" in query:
            speak("okay sir, closing excel")
            os.system("taskkill /f /im EXCEL.EXE")
        
        elif 'open powerpoint' in query or 'open powerpoint presentation' in query or "powerpoint" in query:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
            os.startfile(path)
        
        elif 'close powerpoint' in query or 'close powerpoint presentation' in query or "closeing powerpoint" in query:
            speak("okay sir, closing powerpoint")
            os.system("taskkill /f /im POWERPNT.EXE")

        elif 'calculator' in query or 'open calculator' in query or "open calc" in query or "opening calculator" in query:
            path = "C:\\Windows\\System32\\calc.exe"
            os.startfile(path)

        elif 'close calculator' in query or 'close calc' in query or "closeing calculator" in query:
            speak("okay sir, closing calculator")
            os.system("taskkill /f /im calc.exe")

        elif 'tell me joke' in query or 'tell me a joke' in query or "joke" in query:
            joke = pyjokes.get_joke()
            speak(joke)

        elif 'switch the window' in query or 'change window' in query:
            pyautogui.keyDown("alt")
            pyautogui.press("tab")
            time.sleep(1)
            import time
            pyautogui.keyUp("alt")
        
        elif 'tell me news' in query or 'news' in query:
            speak("please wait sir, fetching the latest news")
            api = "d44049082a9a4f32ad9f322ca5a19331"
            url = f"https://newsapi.org/v2/top-headlines?country=in&apiKey={api}"
            news = requests.get(url).text
            news_dict = json.loads(news)
            speak(news_dict["articles"][0]["title"])
            speak(news_dict["articles"][1]["title"])
            speak(news_dict["articles"][2]["title"])
            speak(news_dict["articles"][3]["title"])
            speak(news_dict["articles"][4]["title"])
            speak(news_dict["articles"][5]["title"])
            speak(news_dict["articles"][6]["title"])
            speak(news_dict["articles"][7]["title"])
            speak(news_dict["articles"][8]["title"])
            speak(news_dict["articles"][9]["title"])
            speak(news_dict["articles"][10]["title"])
            speak("for more news visit: https://tinyurl.com/w2apzt9h")

        elif 'where i am' in query or 'where we are' in query or "current location" in query:
            speak("wait sir, let me check")
            ip_url = "https://get.geojs.io/v1/ip/geo.json"
            geo_q = requests.get(ip_url)
            geo_d = geo_q.json()
            city = geo_d['city']
            country = geo_d['country']
            speak(f"sir i am not sure, but i think we are in {city} city of {country}")

        elif"weather" in query:
            search = "weather in mansa"
            url = f"https://www.google.com/search?q={search}"
            r = requests.get(url)
            data = BeautifulSoup(r.text, "html.parser")
            temp = data.find("div", class_ = "BNeawe").text # type: ignore
            speak(f"current {search} is {temp}")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")

        elif "ip address" in query or "ip" in query:
            ip = requests.get('https://api.ipify.org').text
            speak(f"your IP address is {ip}")
            print(ip)

        elif "open camera" in query or "open webcam" in query or "opening camera" in query: # close esc key 
            cap = cv2.VideoCapture(0)
            while True:
                ret, img = cap.read()
                cv2.imshow('webcam', img)
                k = cv2.waitKey(50)
                if k == 27:
                    break
            cap.release()
            cv2.destroyAllWindows()
        
        elif "battery" in query or "Check battery" in query or "battery percentage" in query:
            battery = psutil.sensors_battery()
            percentage = battery.percent
            speak(f"our battery is at {percentage} percent")

        elif"opening mail" in query or "open mail" in query or "mail" in query:
            webbrowser.open("gmail.com")
            speak("opening mail")

        elif"lock the system" in query or "lock system" in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()
            speak("done sir the system is locked")

        elif "opening task manager" in query:
            speak("opening task manager")
            os.startfile("C:\\Windows\\System32\\taskmgr.exe")
        
        elif ("close task manager" in query):
            speak("closing task manager")
            os.system("taskkill /f /im taskmgr.exe")

        elif "take a picture" in query or "take a photo" in query or "taking a picture" in query or "taking a photo" in query:
            speak("please specify the name of file")
            name = takeCommand()
            cam = cv2.VideoCapture(0)
            speak("please wait while i am taking the picture")
            ret, img = cam.read()
            cv2.imwrite(f"{name}.jpg", img)
            cam.release()
            speak("done sir")
            cv2.destroyAllWindows()
                                                 
        elif "screenshot" in query or "take screenshot" in query:
            img = pyautogui.screenshot()
            speak("taking screenshot")
            speak("please specify the name of file")
            name = takeCommand()
            img.save(f"{name}.png")
            speak("screenshot taken")
            speak("please check your desktop folder")
            os.system("C:\\Users\\Khush\\Desktop\\screenshot.png")
            time.sleep(5)
            speak("done sir")

        
        elif "close all  tabs" in query:
            speak("closing all tabs")
            pyautogui.hotkey("ctrl", "w")

        elif "volume up" in query:
            pyautogui.press("volumeup")

        elif"volume down" in query:
            pyautogui.press("volumedown")

        elif" mute" in query or "mute volume" in query:
            pyautogui.press("mute")

        elif "unmute" in query or "unmute volume" in query:
            pyautogui.press("unmute")

        elif "open google scholar" in query or "opening google scholar" in query or "scholar" in query:
            speak("opening google scholar")
            webbrowser.open("https://scholar.google.com/")

        elif "opening settings" in query or "open settings" in query or "settings" in query or "open system settings" in query:
            speak("opening setting")
            path = "C:\\Windows\\System32\\control.exe"
            os.startfile(path)

        elif "close settings" in query or "close system settings" in query or "closing settings" in query:
            speak("closing settings")
            path = "C:\\Windows\\System32\\control.exe"
            os.system("taskkill /f /im control.exe")

        elif "opening system properties" in query or "open system properties" in query or "system properties" in query:
            speak("opening system properties")
            path = "C:\\Windows\\System32\\sysdm.cpl"
            os.startfile(path)

        elif "close system properties" in query or "close system properties" in query:
            speak("closing system properties")
            path = "C:\\Windows\\System32\\sysdm.cpl"

        elif "sleep" in query or "going to sleep" in query:
            speak("Going to sleep for 15 seconds")
            time.sleep(15)

        elif 'shutdown the system' in query  or "shutdown" in query:
            os.system("shutdown /s /t 5")

        elif 'restart the system' in query or "restart" in query:
            os.system("shutdown /r /t 5")

        elif 'hibernate the system' in query or "hibernate" in query:
            os.system("shutdown /h /t 5")

        elif 'logout' in query or "Signout" in query:
            os.system("shutdown -l")

        elif  "Tell me joke"in query or "joke" in query or "jokes" in query:
            joke = pyjokes.get_joke()
            speak(joke)

        elif "open control panel" in query or "open control panel" in query or "control panel" in query:
            speak("opening control panel")
            path = "C:\\Windows\\System32\\control.exe"
            os.startfile(path)
        
        elif "close control panel" in query or "close control panel" in query:
            speak("closing control panel")
            path = "C:\\Windows\\System32\\control.exe"
            os.system("taskkill /f /im control.exe")
        
        elif "minimize all" in query or "minimize all windows" in query:
            speak("minimizing all windows")
            pyautogui.keyDown("win")
            pyautogui.press("d")
            pyautogui.keyUp("win")

        elif "maximize all" in query or "maximize all windows" in query:
            speak("maximizing all windows")
            pyautogui.keyDown("win")
            pyautogui.press("d")
            pyautogui.keyUp("win")

        elif "news" in query or "news headlines" in query or "news for today" in query:
            speak("opening google news")
            webbrowser.open("https://news.google.com/topstories?hl=en-IN&gl=IN&ceid=IN:en")


        elif "type" in query or "typing" in query or "write" in query or "write something" in query:
            speak("what should i type")
            query = takeCommand()
            speak("typing " + query)
            time.sleep(1)
            pyautogui.write(query)
            
        elif "press tab" in query or "press tab key" in query or "tab" in query:
            pyautogui.press("tab")

        elif "press backspace" in query or "press backspace key" in query or "backspace" in query:
            pyautogui.press("backspace")

        elif "press space" in query or "press space key" in query or "space" in query:
            pyautogui.press("space")

        elif "press shift" in query or "press shift key" in query or "shift" in query:
            pyautogui.press("shift")

        elif "press ctrl" in query or "press ctrl key" in query or "ctrl" in query:
            pyautogui.press("ctrl")

        elif "press alt" in query or "press alt key" in query or "alt" in query:
            pyautogui.press("alt")

        elif "press delete" in query or "press delete key" in query or  "delete" in query:
            pyautogui.press("delete")

        elif "press esc" in query or "press esc key" in query or "esc" in query or "escape" in query:
            pyautogui.press("esc")

        elif "press up" in query or "press up key" in query or "up" in query:
            pyautogui.press("up")

        elif "press down" in query or "press down key" in query or "down" in query:
            pyautogui.press("down")

        elif "press left" in query or "press left key" in query or "left" in query:
            pyautogui.press("left")

        elif "press right" in query or "press right key" in query or "right" in query:
            pyautogui.press("right")

        elif "press f1" in query or "press f1 key" in query or "f1" in query:
            pyautogui.press("f1")

        elif "press f2" in query or "press f2 key" in query or "f2" in query:
            pyautogui.press("f2")

        elif "press f3" in query or "press f3 key" in query or "f3" in query:
            pyautogui.press("f3")

        elif "press f4" in query or "press f4 key" in query or "f4" in query:
            pyautogui.press("f4")

        elif "press f5" in query or "press f5 key" in query or "f5" in query:
            pyautogui.press("f5")

        elif "press f6" in query or "press f6 key" in query or "f6" in query:
            pyautogui.press("f6")

        elif "press f7" in query or "press f7 key" in query or "f7" in query:
            pyautogui.press("f7")

        elif "press f8" in query or "press f8 key" in query or "f8" in query:
            pyautogui.press("f8")

        elif "press f9" in query or "press f9 key" in query or "f9" in query:
            pyautogui.press("f9")

        elif "press f10" in query or "press f10 key" in query or "f10" in query:
            pyautogui.press("f10")

        elif "press f11" in query or "press f11 key" in query or "f11" in query:
            pyautogui.press("f11")

        elif "press f12" in query or "press f12 key" in query or "f12" in query:
            pyautogui.press("f12")

        elif "press print screen" in query or "press print screen key" in query or "print screen" in query:
            pyautogui.press("printscreen")

        elif "press pause" in query or "press pause key" in query or "pause" in query:
            pyautogui.press("pause")

        elif "press insert" in query or "press insert key" in query or "insert" in query:
            pyautogui.press("insert")

        elif "press home" in query or "press home key" in query or "home" in query:
            pyautogui.press("home")

        elif "press end" in query or "press end key" in query or "end" in query:
            pyautogui.press("end")

        elif "press page up" in query or "press page up key" in query or "page up" in query:
            pyautogui.press("pageup")

        elif "press page down" in query or "press page down key" in query or "page down" in query:
            pyautogui.press("pagedown")

        elif "press num lock" in query or "press num lock key" in query  or "num lock" in query:
            pyautogui.press("numlock")

        elif "press caps lock" in query or "press caps lock key" in query or "caps lock" in query:
            pyautogui.press("capslock")

        elif "press scroll lock" in query or "press scroll lock key" in query or "scroll lock" in query:
            pyautogui.press("scrolllock")

        elif "press clear" in query or "press clear key" in query or "clear" in query:
            pyautogui.press("clear")

        elif "press delete" in query or "press delete key" in query or "delete" in query:
            pyautogui.press("delete")

        elif "select all" in query or "select all key" in query or "all" in query:
            pyautogui.hotkey('ctrl', 'a')

        elif "copy" in query or "copy key" in query or "copy" in query:
            pyautogui.hotkey('ctrl', 'c')

        elif "paste" in query or "paste key" in query or "paste" in query:
            pyautogui.hotkey('ctrl', 'v')

        elif "undo" in query or "undo key" in query or "undo" in query:
            pyautogui.hotkey('ctrl', 'z')

        elif "redo" in query or "redo key" in query or "redo" in query:
            pyautogui.hotkey('ctrl', 'y')

        elif "cut" in query or "cut key" in query or "cut" in query:
            pyautogui.hotkey('ctrl', 'x')

        elif "show history" in query or "show history " in query or "history" in query:
            pyautogui.hotkey('ctrl', 'h')

        elif "show find" in query or "show find " in query or "find" in query:
            pyautogui.hotkey('ctrl', 'f')

        elif "open file" in query or "open file key" in query or "file" in query: 
            pyautogui.hotkey('ctrl', 'o')
        
        elif "save file" in query or "save file key" in query or "file" in query: 
            pyautogui.hotkey('ctrl', 's')

        elif "close file" in query or "close file key" in query or "file" in query: 
            pyautogui.hotkey('ctrl', 'w')

        elif "new file" in query or "new file key" in query or "file" in query: 
            pyautogui.hotkey('ctrl', 'n')

        elif "print file" in query or "print file key" in query or "file" in query: 
            pyautogui.hotkey('ctrl', 'p')

        elif "new folder" in query or "new folder key" in query or "folder" in query: 
            pyautogui.hotkey('ctrl', 'shift', 'n')

        elif "press enter " in query or "more enter key" in query or "enter" in query: 
            pyautogui.hotkey('ctrl', 'shift', 'enter')

        elif "open website" in query or "open url" in query or "open site" in query:
            speak("what should i open")
            query = takeCommand().lower()
            webbrowser.open(f"{query}.com")

        elif "close window" in query:
            speak("Okay")
            pyautogui.hotkey('alt', 'f4')
        
        elif "stop " in query or " Stop" in query:
            speak("Okay")
            break

        elif  "exit" in query or "quit" in query or "close" in query:
            speak("Thanks for using me.")
            exit()


        else:
            speak("Sorry, I didn't get that")