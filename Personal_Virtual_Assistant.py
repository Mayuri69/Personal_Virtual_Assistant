import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import pywhatkit
import os
import smtplib
import pyjokes
import subprocess
import ctypes
import json
import winshell
import requests
import win32com.client as wincl
from urllib.request import urlopen
import shutil
import time


from ecapture import ecapture as ec

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')

engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

    speak("I am your personal assistant Madam. Please tell me how may I help you ")


def takeCommand():

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:

        print("Say that again please...")
        speak("Say that again please...")
        return "None"
    return query


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('your email address', 'password')
    server.sendmail('your email', to, content)
    server.close()


if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand().lower()

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube music' in query:
            speak("Here you go to Youtube Music\n")
            webbrowser.open("music.youtube.com")

        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")

        elif 'play' in query:
            song = query.replace("play", " ")
            speak("Playing " + song)
            pywhatkit.playonyt(song)

        elif ' send whats app message' in query:
            pywhatkit.sendwhatmsg("+9136353841","Hello from Mayuri",22, 54)
            print("Succesfully sent!")

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'open facebook' in query:
            speak("Here you go to Facebook\n")
            webbrowser.open("facebook.com")

        elif 'open amazon' in query:
            speak("Here you go to Amazon\n")
            webbrowser.open("amazon.in")

        elif 'open flipkart' in query:
            speak("Here you go to Flipkart\n")
            webbrowser.open("flipkart.com")

        elif 'open gmail' in query:
            speak("Here you go to Gmail\n")
            webbrowser.open("gmail.com")

        elif 'play music' in query:
            speak("Here you go with music")
            music_dir = 'E:\music'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[0]))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'search' in query or 'play' in query:

            query = query.replace("search", "")
            query = query.replace("play", "")
            webbrowser.open(query)

        elif 'open chrome' in query:
            speak("Here you go to Chrome\n")
            codePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            os.startfile(codePath)

        elif 'open python' in query:
            speak("Here you go to Python\n")
            codePath = "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2019.3.1\\bin\\pycharm64.exe"
            os.startfile(codePath)

        elif 'open sublime' in query:
            speak("Here you go to Sublime\n")
            codePath = "D:\Sublime Text 3\sublime_text.exe"
            os.startfile(codePath)

        elif 'open amazon kindle' in query:
            speak("Here you go to Amazon Kindle\n")
            codePath = "C:\\Users\\JAXZ\\AppData\\Local\\Amazon\\Kindle\\application\\Kindle.exe"
            os.startfile(codePath)

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif 'email to Mayuri' in query or 'send a mail to Mayuri' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "mayurithorve69@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry sir. I am not able to send this email")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'fine' in query or "mast" in query:
            speak("It's good to know that your fine")

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Mayuri .")

        elif "who i am" in query or "main kaun hun" in query:
            speak("If you talk then definately your human.")

        elif 'what is love' in query or " what's love" in query:
            speak("It is 7th sense that destroy all other senses")

        elif "who are you" in query:
            speak("I am your Personal assistant created by Mayuri ")

        elif 'reason' in query:
            speak("I was created as a Minor project by Miss Mayuri")

        elif "why you came to world" in query:
            speak("Thanks to Mayuri. further It's a secret")

        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign- out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('data.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
        elif "show note" in query or "show notes" in query or "show me my notes" in query:
            speak("Showing Notes")
            file = open("data.txt", "r")
            print(file.read())
            speak(file.read(6))

        elif "will you be my gf" in query or "will you be my bf" in query:
            speak("I'm not sure about, may be you should give me some time")

        elif "good morning" in query:
            speak("A warm" + query)
            speak("How are you Sir")

        elif "afternoon" in query:
            speak(query + "Sir")
            speak("How are you Sir")

        elif "evening" in query:
            speak(query + "Sir")
            speak("How are you Sir")

        elif "night" in query:
            speak(query + "Sir")
            speak("Have a good sleep")

        elif "i love you" in query:
            speak("It's hard to understand")

        elif "what's you doing" in query or "what are you doing" in query:
            speak("Being with you,if you want me around")

        elif "bore" in query or "feeling bored" in query:
            speak("Alright, I can you a joke then")

        elif 'news' in query:

            try:
                jsonObj = urlopen('''http://newsapi.org/v2/everything?domains=wsj.com&apiKey=83d3a7ff48f54250805df0302fb6eecf''')
                data = json.load(jsonObj)
                i = 1
                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============''' + '\n')

                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                print(str(e))

        elif "find" in query:

            app_id = "2URK6A-TWU5XQ4XTP"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('find')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif "calculate" in query:
            app_id = "2URK6A-TWU5XQ4XTP"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop me from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,0,"E:\wallpaper",0)
            speak("Background changed succesfully")

