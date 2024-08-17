import ctypes
import json
import os
import requests
import speech_recognition as sr
import datetime
import pyttsx3
import wikipedia
import webbrowser
import pywhatkit
import pyjokes
import psutil
import tkinter as tk
from tkinter import messagebox, scrolledtext
import threading
import pyautogui
import time
from bs4 import BeautifulSoup

# Initialize text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        query = r.recognize_google(audio, language='en-in')
    except Exception:
        return "none"
    return query.lower()

def wish():
    hour = int(datetime.datetime.now().hour)
    if hour < 12:
        speak("Good Morning sir")
    elif hour < 18:
        speak("Good Afternoon sir")
    else:
        speak("Good Evening sir")
    speak("I am Personal Assistant. How may I help you?")

def run_voice_commands():
    wish()
    while running:
        query = takeCommand()
        if query == "none":
            continue
        
        if "jarvis" in query or "alexa" in query:
            query = query.replace("jarvis", "")
            if 'wikipedia' in query:
                query = query.replace("wikipedia", "")
                results = wikipedia.summary(query, sentences=2)
                speak(results)
                display_result(results)
            
            elif 'open notepad' in query:
                os.startfile("C:\\Windows\\System32\\notepad.exe")
                display_result("Opening Notepad")
            
            elif 'close notepad' in query:
                os.system("taskkill /f /im notepad.exe")
                display_result("Closing Notepad")
            
            elif 'open command prompt' in query:
                os.system("start cmd")
                display_result("Opening Command Prompt")
            
            elif 'close command prompt' in query:
                os.system("taskkill /f /im cmd.exe")
                display_result("Closing Command Prompt")
            
            elif 'open code' in query or 'open visual studio code' in query:
                os.startfile("C:\\Users\\khush\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe")
                display_result("Opening Visual Studio Code")
            
            elif 'close code' in query:
                os.system("taskkill /f /im Code.exe")
                display_result("Closing Visual Studio Code")
            
            elif 'open youtube' in query:
                webbrowser.open("youtube.com")
                display_result("Opening YouTube")
            
            elif 'close youtube' in query:
                os.system("taskkill /f /im msedge.exe")
                display_result("Closing YouTube")
            
            elif 'search on google' in query:
                query = query.replace("search on google", "").strip()
                webbrowser.open(f"https://www.google.com/search?q={query}")
                display_result(f"Searching Google for {query}")
            
            elif 'play on youtube' in query:
                query = query.replace("play on youtube", "").strip()
                pywhatkit.playonyt(query)
                display_result(f"Playing {query} on YouTube")
            
            elif 'the time' in query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"The time is {strTime}")
                display_result(f"The time is {strTime}")
            
            elif 'tell me news' in query or 'news' in query:
                speak("please wait sir, fetching the latest news")
                api = "d44049082a9a4f32ad9f322ca5a19331"
                url = f"https://newsapi.org/v2/top-headlines?country=in&apiKey={api}"
                news = requests.get(url).text
                news_dict = json.loads(news)
                news_list = [news_dict["articles"][i]["title"] for i in range(10)]
                for headline in news_list:
                    speak(headline)
                    display_result(headline)
                display_result("For more news visit: https://tinyurl.com/w2apzt9h")
            
            elif 'battery' in query:
                battery = psutil.sensors_battery()
                percentage = battery.percent
                speak(f"Battery is at {percentage} percent")
                display_result(f"Battery is at {percentage} percent")

            elif"lock the system" in query or "lock system" in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
                speak("done sir the system is locked")
                display_result("done sir the system is locked")

            elif "opening task manager" in query:
                speak("opening task manager")
                os.startfile("C:\\Windows\\System32\\taskmgr.exe")
                display_result("opening task manager")
        
            elif ("close task manager" in query):
                speak("closing task manager")
                os.system("taskkill /f /im taskmgr.exe")
                display_result("closing task manager")

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
                display_result("done sir")


            elif "close all  tabs" in query:
                speak("closing all tabs")
                pyautogui.hotkey("ctrl", "w")
                display_result("closing all tabs")

            elif "volume up" in query:
                pyautogui.press("volumeup")
                display_result("volume up")

            elif"volume down" in query:
                pyautogui.press("volumedown")
                display_result("volume down")

            elif" mute" in query or "mute volume" in query:
                pyautogui.press("mute")
                display_result("mute")

            elif "unmute" in query or "unmute volume" in query:
                pyautogui.press("unmute")
                display_result("unmute")

            elif "open google scholar" in query or "opening google scholar" in query or "scholar" in query:
                speak("opening google scholar")
                webbrowser.open("https://scholar.google.com/")
                display_result("opening google scholar")

            elif "opening settings" in query or "open settings" in query or "settings" in query or "open system settings" in query:
                speak("opening setting")
                path = "C:\\Windows\\System32\\control.exe"
                os.startfile(path)
                display_result("opening settings")

            elif "close settings" in query or "close system settings" in query or "closing settings" in query:
                speak("closing settings")
                path = "C:\\Windows\\System32\\control.exe"
                os.system("taskkill /f /im control.exe")
                display_result("closing settings")

            elif "opening system properties" in query or "open system properties" in query or "system properties" in query:
                speak("opening system properties")
                path = "C:\\Windows\\System32\\sysdm.cpl"
                os.startfile(path)
                display_result("opening system properties")

            elif "close system properties" in query or "close system properties" in query:
                speak("closing system properties")
                path = "C:\\Windows\\System32\\sysdm.cpl"
                os.system("taskkill /f /im sysdm.cpl")
                display_result("closing system properties")

            elif "sleep" in query or "going to sleep" in query:
                speak("Going to sleep for 15 seconds")
                time.sleep(15)
                display_result("Going to sleep for 15 seconds")

            elif 'shutdown the system' in query  or "shutdown" in query:
                os.system("shutdown /s /t 5")
                display_result("Shutting down the system")

            elif 'restart the system' in query or "restart" in query:
                os.system("shutdown /r /t 5")
                display_result("Restarting the system")

            elif 'hibernate the system' in query or "hibernate" in query:
                os.system("shutdown /h /t 5")
                display_result("Hibernating the system")

            elif 'logout' in query or "Signout" in query:
                os.system("shutdown -l")
                display_result("Signing out")

            elif "open control panel" in query or "open control panel" in query or "control panel" in query:
                speak("opening control panel")
                path = "C:\\Windows\\System32\\control.exe"
                os.startfile(path)
                display_result("opening control panel")
        
            elif "close control panel" in query or "close control panel" in query:
                speak("closing control panel")
                path = "C:\\Windows\\System32\\control.exe"
                os.system("taskkill /f /im control.exe")
                display_result("closing control panel")
        
            elif "minimize all" in query or "minimize all windows" in query:
                speak("minimizing all windows")
                pyautogui.keyDown("win")
                pyautogui.press("d")
                pyautogui.keyUp("win")
                display_result("minimizing all windows")

            elif "maximize all" in query or "maximize all windows" in query:
                speak("maximizing all windows")
                pyautogui.keyDown("win")
                pyautogui.press("d")
                pyautogui.keyUp("win")
                display_result("maximizing all windows")

            elif "news" in query or "news headlines" in query or "news for today" in query:
                speak("opening google news")
                webbrowser.open("https://news.google.com/topstories?hl=en-IN&gl=IN&ceid=IN:en")
                display_result("opening google news")


            elif "type" in query or "typing" in query or "write" in query or "write something" in query:
                speak("what should i type")
                query = takeCommand()
                speak("typing " + query)
                time.sleep(1)
                pyautogui.write(query)
                display_result("typing " + query)
            
            elif "press tab" in query or "press tab key" in query or "tab" in query:
                pyautogui.press("tab")
                display_result("pressing tab")

            elif "press backspace" in query or "press backspace key" in query or "backspace" in query:
                pyautogui.press("backspace")
                display_result("pressing backspace")

            elif "press space" in query or "press space key" in query or "space" in query:
                pyautogui.press("space")
                display_result("pressing space")

            elif "press shift" in query or "press shift key" in query or "shift" in query:
                pyautogui.press("shift")
                display_result("pressing shift")

            elif "press ctrl" in query or "press ctrl key" in query or "ctrl" in query:
                pyautogui.press("ctrl")
                display_result("pressing ctrl")

            elif "press alt" in query or "press alt key" in query or "alt" in query:
                pyautogui.press("alt")
                display_result("pressing alt")

            elif "press delete" in query or "press delete key" in query or  "delete" in query:
                pyautogui.press("delete")
                display_result("pressing delete")

            elif "press esc" in query or "press esc key" in query or "esc" in query:
                pyautogui.press("esc")
                display_result("pressing esc")

            elif "press up" in query or "press up key" in query or "up" in query:
                pyautogui.press("up")
                display_result("pressing up")

            elif "press down" in query or "press down key" in query or "down" in query:
                pyautogui.press("down")
                display_result("pressing down")

            elif "press left" in query or "press left key" in query or "left" in query:
                pyautogui.press("left")
                display_result("pressing left")

            elif "press right" in query or "press right key" in query or "right" in query:
                pyautogui.press("right")
                display_result("pressing right")

            elif "press f1" in query or "press f1 key" in query or "f1" in query:
                pyautogui.press("f1")
                display_result("pressing f1")

            elif "press f2" in query or "press f2 key" in query or "f2" in query:
                pyautogui.press("f2")
                display_result("pressing f2")

            elif "press f3" in query or "press f3 key" in query or "f3" in query:
                pyautogui.press("f3")
                display_result("pressing f3")

            elif "press f4" in query or "press f4 key" in query or "f4" in query:
                pyautogui.press("f4")
                display_result("pressing f4")

            elif "press f5" in query or "press f5 key" in query or "f5" in query:
                pyautogui.press("f5")
                display_result("pressing f5")

            elif "press f6" in query or "press f6 key" in query or "f6" in query:
                pyautogui.press("f6")
                display_result("pressing f6")

            elif "press f7" in query or "press f7 key" in query or "f7" in query:
                pyautogui.press("f7")
                display_result  ("pressing f7")

            elif "press f8" in query or "press f8 key" in query or "f8" in query:
                pyautogui.press("f8")
                display_result("pressing f8")

            elif "press f9" in query or "press f9 key" in query or "f9" in query:
                pyautogui.press("f9")
                display_result("pressing f9")

            elif "press f10" in query or "press f10 key" in query or "f10" in query:
                pyautogui.press("f10")
                display_result ("pressing f10")

            elif "press f11" in query or "press f11 key" in query or "f11" in query:
                pyautogui.press("f11")
                display_result("pressing f11")

            elif "press f12" in query or "press f12 key" in query or "f12" in query:
                pyautogui.press("f12")
                display_result("pressing f12")

            elif "press print screen" in query or "press print screen key" in query or "print screen" in query:
                pyautogui.press("printscreen")
                display_result("pressing print screen")

            elif "press pause" in query or "press pause key" in query or "pause" in query:
                pyautogui.press("pause")
                display_result("pressing pause")

            elif "press insert" in query or "press insert key" in query or "insert" in query:
                pyautogui.press("insert")
                display_result("pressing insert")

            elif "press home" in query or "press home key" in query or "home" in query:
                pyautogui.press("home")
                display_result("pressing home")

            elif "press end" in query or "press end key" in query or "end" in query:
                pyautogui.press("end")
                display_result("pressing end")

            elif "press page up" in query or "press page up key" in query or "page up" in query:
                pyautogui.press("pageup")
                display_result("pressing page up")

            elif "press page down" in query or "press page down key" in query or "page down" in query:
                pyautogui.press("pagedown")
                display_result("pressing page down")

            elif "press num lock" in query or "press num lock key" in query  or "num lock" in query:
                pyautogui.press("numlock")
                display_result("pressing num lock")

            elif "press caps lock" in query or "press caps lock key" in query or "caps lock" in query:
                pyautogui.press("capslock")
                display_result("pressing caps lock")

            elif "press scroll lock" in query or "press scroll lock key" in query or "scroll lock" in query:
                pyautogui.press("scrolllock")
                display_result("pressing scroll lock")

            elif "press clear" in query or "press clear key" in query or "clear" in query:
                pyautogui.press("clear")
                display_result("pressing clear")

            elif "press delete" in query or "press delete key" in query or "delete" in query:
                pyautogui.press("delete")
                display_result("pressing delete")

            elif "select all" in query or "select all key" in query or "all" in query:
                pyautogui.hotkey('ctrl', 'a')
                display_result("selecting all")

            elif "copy" in query or "copy key" in query or "copy" in query:
                pyautogui.hotkey('ctrl', 'c')
                display_result ("copying")

            elif "paste" in query or "paste key" in query or "paste" in query:
                pyautogui.hotkey('ctrl', 'v')
                display_result("pasting")

            elif "undo" in query or "undo key" in query or "undo" in query:
                pyautogui.hotkey('ctrl', 'z')
                display_result("undoing")

            elif "redo" in query or "redo key" in query or "redo" in query:
                pyautogui.hotkey('ctrl', 'y')
                display_result("redoing")

            elif "cut" in query or "cut key" in query or "cut" in query:
                pyautogui.hotkey('ctrl', 'x')
                display_result("cutting")

            elif "show history" in query or "show history " in query or "history" in query:
                pyautogui.hotkey('ctrl', 'h')
                display_result("showing history")

            elif "show find" in query or "show find " in query or "find" in query:
                pyautogui.hotkey('ctrl', 'f')
                display_result("showing find")

            elif "open file" in query or "open file key" in query or "file" in query: 
                pyautogui.hotkey('ctrl', 'o')
                display_result("opening file")
        
            elif "save file" in query or "save file key" in query or "file" in query: 
                pyautogui.hotkey('ctrl', 's')
                display_result("saving file")

            elif "close file" in query or "close file key" in query or "file" in query: 
                pyautogui.hotkey('ctrl', 'w')
                display_result("closing file")

            elif "new file" in query or "new file key" in query or "file" in query: 
                pyautogui.hotkey('ctrl', 'n')
                display_result ("new file")

            elif "print file" in query or "print file key" in query or "file" in query: 
                pyautogui.hotkey('ctrl', 'p')
                display_result("print file")

            elif "new folder" in query or "new folder key" in query or "folder" in query: 
                pyautogui.hotkey('ctrl', 'shift', 'n')
                display_result("new folder")

            elif "press enter " in query or "more enter key" in query or "enter" in query: 
                pyautogui.hotkey('ctrl', 'shift', 'enter')
                display_result("pressing enter")

            elif "open website" in query or "open url" in query or "open site" in query:
                speak("what should i open")
                query = takeCommand().lower()
                webbrowser.open(f"{query}.com")
                display_result(f"Opening {query}")

            elif "close window" in query:
                speak("Okay")
                pyautogui.hotkey('alt', 'f4')
                display_result("Closing window")


            elif 'open word' in query or 'open word document' in query or "word" in query:
                path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
                os.startfile(path)
                display_result("opening word")

            elif 'close word' in query or 'close word document' in query or "close word" in query:
                speak("okay sir, closing word")
                os.system("taskkill /f /im WINWORD.EXE")
                display_result("closing word")

            elif 'open excel' in query or 'open excel sheet' in query or "opeining excel" in query:
                path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
                os.startfile(path)
                display_result("opening excel")

            elif 'close excel' in query or 'close excel sheet' in query or "closeing excel" in query:
                speak("okay sir, closing excel")
                os.system("taskkill /f /im EXCEL.EXE")
                display_result("closing excel")
        
            elif 'open powerpoint' in query or 'open powerpoint presentation' in query or "powerpoint" in query:
                path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
                os.startfile(path)
                display_result("opening powerpoint")
        
            elif 'close powerpoint' in query or 'close powerpoint presentation' in query or "closeing powerpoint" in query:
                speak("okay sir, closing powerpoint")
                os.system("taskkill /f /im POWERPNT.EXE")
                display_result("closing powerpoint")

            elif 'calculator' in query or 'open calculator' in query or "open calc" in query or "opening calculator" in query:
                path = "C:\\Windows\\System32\\calc.exe"
                os.startfile(path)
                display_result("opening calculator")

            elif 'close calculator' in query or 'close calc' in query or "closeing calculator" in query:
                speak("okay sir, closing calculator")
                os.system("taskkill /f /im calc.exe")
                display_result("closing calculator")

            elif 'switch the window' in query or 'change window' in query:
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                time.sleep(1)
                import time
                pyautogui.keyUp("alt")
                display_result("switching the window")

            elif 'how are you' in query:
                speak("I am fine, Thank you")
                speak("How are you, Sir")
                display_result("I am fine, Thank you")

            elif 'fine' in query or "good" in query:
                speak("It's good to know that your fine")
                display_result("It's good to know that your fine")

            elif "ip address" in query or "ip" in query:
                ip = requests.get('https://api.ipify.org').text
                speak(f"your IP address is {ip}")
                print(ip)
                display_result(f"your IP address is {ip}")

            elif"weather" in query:
                search = "weather in mansa"
                url = f"https://www.google.com/search?q={search}"
                r = requests.get(url)
                data = BeautifulSoup(r.text, "html.parser")
                temp = data.find("div", class_ = "BNeawe").text # type: ignore
                speak(f"current {search} is {temp}")
        
            elif "stop " in query or " Stop" in query:
                speak("Okay")
                break
                display_result("Okay")
            
            elif 'quit' in query or 'exit' in query or "close" in query:
                speak("Goodbye sir. Have a nice day!")
                display_result("Goodbye sir. Have a nice day!")
                stop_voice_commands()


        else:
            speak("I didn't get that. Please say it again.")
            display_result("I didn't get that. Please say it again.")

def display_result(text):
    result_text.config(state=tk.NORMAL)
    result_text.insert(tk.END, f"{text}\n")
    result_text.config(state=tk.DISABLED)
    result_text.yview(tk.END)

def start_stop():
    global running
    if running:
        stop_voice_commands()
    else:
        running = True
        start_button.config(text="Stop", bg="lightcoral")
        thread = threading.Thread(target=run_voice_commands)
        thread.start()

def stop_voice_commands():
    global running
    running = False
    start_button.config(text="Start", bg="lightgreen")

def show_commands():
    commands = """Commands:
    - Open/Close Notepad
    - Open/Close code
    - Open/Close Command Prompt
    - Open/Close Visual Studio Code
    - Search on Google
    - Play on YouTube
    - Get the current time
    - Get news headlines
    - Check battery percentage
    - Lock the system
    - Open Task Manager
    - Take Screenshot
    - Show history
    - Show find
    - Open file
    - Save file
    - Close file
    - New file
    - Print file
    - New folder
    - Press enter
    - Open website
    - Stop
    - close window
    - minimize window/ maximize window
    - type
    - press tab, backspace, space, shift, alt, ctrl, enter, delete, escape, up, down, left, right, page up, page down, home, end, insert, f1 to f12 and f keys , print screen, scroll lock, pause, num lock, caps lock
    - selet all, copy, cut, paste, undo, redo
    - open file
    - save file
    - close file
    - new file
    - print file
    - new folder
    - press enter
    - close window 
    - open website
    - open /close word
    - open /close excel
    - open /close powerpoint
    - open /close calculator
    - switch the window
    - how are you
    - ip address
    - weather
    - stop
    - Quit/Exit
    """
    messagebox.showinfo("Command Details", commands)

def on_enter(event):
    event.widget.config(bg="lightblue")

def on_leave(event):
    event.widget.config(bg="SystemButtonFace")

# Set up GUI
root = tk.Tk()
root.title("Voice Assistant")

running = False

start_button = tk.Button(root, text="Start", command=start_stop, bg="lightgreen")
start_button.pack(pady=10)
start_button.bind("<Enter>", on_enter)
start_button.bind("<Leave>", on_leave)

command_details_button = tk.Button(root, text="Command Details", command=show_commands, bg="lightblue")
command_details_button.pack(pady=10)
command_details_button.bind("<Enter>", on_enter)
command_details_button.bind("<Leave>", on_leave)

result_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=15, width=50, state=tk.DISABLED)
result_text.pack(pady=10)

root.mainloop()
