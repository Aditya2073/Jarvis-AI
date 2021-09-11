# Modules that need to be installed
import keyboard
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import pyjokes
import os
import requests

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')

engine.setProperty('voice', voices[0].id)     # Setting voices [0] for male voice and [1] for female voice

# Speak function
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

# Function that wishes the user
def wishme():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        print("Good Morning !")
        speak("Good Morning !")

    elif 12 <= hour < 18:
        print("Good Afternoon !")
        speak("Good Afternoon !")

    else:
        print("Good Evening !")
        speak("Good Evening !")

    print("I am Jarvis. Please tell me what can i do for you.")
    speak("I am Jarvis. Please tell me what can i do for you.")

# voice recognition function
def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        r.energy_threshold = 600
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio)
        print(f"User said: {query}\n")

    except Exception as e:
        print("Please Say it again...")
        return "None"
    return query

# Function to bring the Latest news
def NewsFromBBC():
    # BBC news api
    # Internet connection required to fech the NEWS.....
    main_url = " https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey=4dbc17e007ab436fb66416009dfb59a8"

    open_bbc_page = requests.get(main_url).json()

    # getting all articles in a string article

    article = open_bbc_page["articles"]

    # empty list which will

    # contain all trending news
    results = []

    for ar in article:
        results.append(ar["title"])

    for i in range(len(results)):
        # printing all trending news

        print(i + 1, results[i])

        # to read the news out loud for user

    from win32com.client import Dispatch

    speak = Dispatch("SAPI.Spvoice")

    speak.Speak(results)

# Function To get jokes
def jo():
    pg = pyjokes.get_joke()
    print(pg)
    speak(pg)


if __name__ == '__main__':
    wishme()
    while True:
    # if 1:
        query = take_command().lower()

        if 'wikipedia' in query:   # Search Wikipedia
            try:
                speak("Searching Wikipedia...")
                query = query.replace("wikipedia", "")
                results = wikipedia.summary(query, sentences=2)
                speak("According to wikipedia")
                print(results)
                speak(results)
            except Exception as e:
                print("Could not found you requested..  :(")



        elif 'open youtube' in query:    # Open Youtube
            webbrowser.open("youtube.com")

        elif 'open google' in query:     # Open Google
            webbrowser.open("google.com")

        elif 'open instagram' in query:  # Open Instagram
            webbrowser.open("") # Enter the loged in URL of your instagram account

        elif 'play music' in query:   # Play Music
            music_dir = 'C:\\Users\\Aditya\\Music' # Enter your Music Directory
            songs = os.listdir(music_dir)
            print(songs[0])
            os.startfile(os.path.join(music_dir, songs[0]))
            print("Pres Enter to continue..")
            key = keyboard.read_key()
            if key == 'enter':
                continue

        elif 'the time' in query:  # Get the time
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"sir, The time is {strTime}")

        elif 'open pycharm' in query:  # This elif is optional. If not need comment this elif!!
            charmpath = "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2020.1\\bin\\pycharm64.exe"
            os.startfile(charmpath)

        elif 'latest news' in query:  # To read the news
            NewsFromBBC()

        elif 'tell me jokes' in query:  # To read the Jokes
            jo()
