
#*The subprocess module allows spawning of new processes, connecting to their input/output/error pipes, and obtaining their return codes.
import subprocess
#*pyttsx3 is a python text to speech library.
import pyttsx3
#*tkinter is a gui creator for python
import tkinter
#*json module is a javascrpit object decoder for python
import json
#*random module for giving random results
import random
#*Python Arthimetic operators for math
import operator
#*Python Speech Recognition library
import speech_recognition as sr
#*datetime module to give the date and time
import datetime
#*wikipedia module to give defenitions
import wikipedia
#*webbrowser module to open links
import webbrowser
import os
import winshell
#*python jokes
import pyjokes
#*module to bring rss feeds for news
import feedparser
#*smpt module to send mails
import smtplib
#*binds c language data types
import ctypes
#*time module to give time
import time
#*https requests module to receive data from websites
import requests
#*module to edit folder data
import shutil
#*twilio module to connect to twilio API, used to receive and send sms, voice calls, etc
from twilio.rest import Client
#*Python web scraper
from bs4 import BeautifulSoup
#*module for windows application automation
import win32com.client as wincl
#*module to open links
from urllib.request import urlopen
#*turtle module for gui related commands
from turtle import *
#*freegames module for simple games
from freegames import line
#*Coingecko API wrapper for python to get cryptocurrency prices
import pycoingecko
#*Coingecko API https get requester
from pycoingecko import CoinGeckoAPI

#*browser settings for opening links
chrome = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
edge = 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe %s'
brave = 'C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe %s'

#*change edge to your main browser
browser = edge

#*sets var cg for easier execution of crypto related commands
cg = CoinGeckoAPI()
#*default currency setting for crypto prices
currency = 'usd'

#*Sets tts settings such as output voice and input type
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
#*sets output voice to male or female. voices[1] is for female, whereas it can be changed to voices[0] for male
engine.setProperty("voice", voices[1].id)
#*sets a name for the assistant
assistantname = ("Google")

#*function for voice input 
def speak(audio=""):
    engine.say(audio)
    engine.runAndWait()

#*Says greetings upon bot startup
def wishMe():
    #*takes the current hour using datetime module
    hour = int(datetime.datetime.now().hour)
    
    #*if the received time is from 12:01AM to 11:59AM it says and prints "Good Morning"
    if hour >= 0 and hour < 12:
        print("Good Morning")
        speak("Good Morning")
        
    #*if the received time is from 12:00pm to 5:59pm it says and prints "Good Afternoon"
    elif hour >= 12 and hour < 18:
        print("Good Afternoon")
        speak("Good Afternoon")
        
    #*if the received time is from 6:00pm to 7:59pm it says and prints "Good Evening"
    elif hour >= 18 and hour < 20:
        print("Good Evening")
        speak("Good Evening")
       
    #*if it received any of the times not mentioned above, it just prints "Hello"  
    else:
        print("Hello")
        speak("Hello")
    
    #*Says and prints it's name
    print("I am Your Assistant, " + assistantname)
    speak("I Am Your Assistant, " + assistantname)

#*the function used to take and process voice command
def takeCommand():
     
     #*uses sr module to check if there is any input  
    r = sr.Recognizer()
     
    #*takes device microphone as input
    with sr.Microphone() as source:
        
        #*print's "Listening..." until it receives audio, and wait's 1 second after audio
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
  
    try:
        #*uses google speech recognition API to register query
        print("Recognizing...")   
        query = r.recognize_google(audio, language ='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        #*if no voice is heard or is not able to understand, it prints "Unable to Recognize your voice" 
        print(e)   
        print("Unable to Recognize your voice.") 
        return "None"
     
    return query

def connect4():
    turns = {'red': 'yellow', 'yellow': 'red'}
    state = {'player': 'yellow', 'rows': [0] * 8}

    def grid():
        "Draw Connect Four grid."
        bgcolor('light blue')

        for x in range(-150, 200, 50):
            line(x, -200, x, 200)

        for x in range(-175, 200, 50):
            for y in range(-175, 200, 50):
                up()
                goto(x, y)
                dot(40, 'white')

        update()

    def tap(x, y):
        "Draw red or yellow circle in tapped row."
        player = state['player']
        rows = state['rows']

        row = int((x + 200) // 50)
        count = rows[row]

        x = ((x + 200) // 50) * 50 - 200 + 25
        y = count * 50 - 200 + 25

        up()
        goto(x, y)
        dot(40, player)
        update()

        rows[row] = count + 1
        state['player'] = turns[player]

    setup(420, 420, 370, 0)
    hideturtle()
    tracer(False)
    grid()
    onscreenclick(tap)
    done()

#*The sendMail function below is still underprogress
'''
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
     
    #* Enable low security in gmail
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()
'''

#*clears previous text from terminal and executes the Greeting function
if __name__ == '__main__':
    clear = lambda: os.system('cls')

    clear()
    wishMe()
        
#*once query is registered it is registered and processed here to give an output
while True:
        
        #*query is moved to all lowercase for easier recognition by the API
        query = takeCommand().lower()
  #!#--------------------------------------------Link Opening Commands-------------------------------------------- #!#        

        #*if query strictly says to open google, it is done
        if query == 'open google':
            print("Opening Google...\n")
            speak("Opening Google.com\n")
            webbrowser.get(browser).open("google.com")
    
        #*searches google
        elif "search for" in query or "google" in query:
            query = query.replace("search", "")
            query = query.replace("search for", "")
            query = query.replace("google", "")
            query = query.replace("on chrome", "")
            query = query.replace("on google", "")
            print("Googling " + query)
            speak("Googling " + query)
            webbrowser.get(browser).open("https://www.google.com/search?q=" + query)

        #*if wikipedia is heard in the query, it will search wikipedia for the term given 
        elif 'wikipedia' in query:
            print("Searching Wikipedia...")
            speak('Searching Wikipedia...')
            query = query.replace("search wikipedia for", "")
            query = query.replace("search wikipedia", "")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences = 3)
            print("According to Wikipedia...\n")
            speak("According to Wikipedia")
            print(results)
            speak(results)

        #*if query strictly says to open youtube, it is done
        elif query == 'open youtube':
            print("Opening Youtube...\n")
            speak("Opening Youtube...\n")
            webbrowser.get(browser).open("youtube.com")
            
        elif "play" in query and "on youtube" in query:
            query = query.replace ("play", "")
            query = query.replace ("on youtube", "")
            print("Searching for " + query + "on Youtube")
            speak("Searching for " + query + "on Youtube")
            search = query
            webbrowser.get(browser).open("https://www.youtube.com/results?search_query=" + search)
            
        elif "play" in query and "on spotify" in query:
            query = query.replace("play", "")
            query = query.replace("on spotify", "")
            print("Searching Spotify for " + query)
            speak("Searching spotify for " + query)
            search = query
            webbrowser.get(browser).open("https://open.spotify.com/search/" + search)

        elif "play" in query and ("on gaana" or "on ganna" or "on gana")in query:
            query = query.replace("play", "")
            query = query.replace("on gaana", "")
            print("Searching Gaana for: " + query)
            speak("Searching Gaana for: " + query)
            search = query
            webbrowser.get(browser).open("https://gaana.com/search/" + search)
            
        #*uses webbrowser modules to open a maps.google.com site
        elif "where is" in query:
            query = query.replace("where is", "")
            print("Searching for " + query + " on Google Maps")
            speak("Searching for " + query + " on Google Maps")
            location = query
            webbrowser.open("https://www.google.com/maps/place/" + location)
        
        
 #!#--------------------------------------------General Questions-------------------------------------------- #!#
        
        
         #*if 'the time' is heard in query, it will display time using datetime module    
        elif 'the time' in query:
            now = datetime.now()
            strTime = now.strftime("%H:%M.%S")
            print(f"The time is {strTime}")
            Time = now.strftime("%H:%M")
            speak(f"The time is {Time}")
            
        #*if 'how are you' is heard in query, it says taht it is fine
        elif 'how are you' in query or "how are you" in query or "hru" in query or "how r u" in query:
            print("I am fine, Thank You")
            speak("I am fine, Thank You")
        
        #*if anything is heard about it's name, it uses the assistantname var defined earlier
        elif "what's your name" in query or "what is your name" in query or "whats your name in query":
            print("My friends call me", assistantname)
            speak("My friends call me " + assistantname)
            
               
 #!#--------------------------------------------Cryptocurrency Commands-------------------------------------------- #!#
            
            
        elif "price" in query:
            print("What crypto currency would you like to find the price of?")
            speak("What crypto currency would you like to find the price of?")
            r = sr.Recognizer()
            with sr.Microphone() as source:
                r.pause_threshold = 1
                audio = r.listen(source)
  
            try:
                query = r.recognize_google(audio, language ='en-in')
            
            except Exception as e:
                print("Unable to Recognize your voice.")
            
            search = query.lower()
            
            try:
                cryptoprice = cg.get_price(ids=query, vs_currencies=currency)
                price = cryptoprice[query]  
                print(query + " is currently worth " + str(price[currency]) + " " + currency)
                speak(query + " is currently worth " + str(price[currency]) + " " + currency)
            except Exception as e:
                print(e)
                print("Sorry I couldn't find that currency on Coingecko")
                speak("Sorry I couldn't find that currency on Coingecko")
                

 #!#--------------------------------------------Fun Commands-------------------------------------------- #!#
        
        
        #*if joke is heard, it gives a python joke
        elif 'python joke' in query or "tell me a joke" in query:
            joke = pyjokes.get_joke()
            print(joke)
            speak(joke)
            
        elif "connect4" in query or "connect 4" in query or "connect four" in query:
            print("Opening Connect 4")
            speak("Opening Connect Four")
            speak("Connect 4 is now open. If you can't see it, please check for a new window")
            connect4()
                
 #!#--------------------------------------------Device Control-------------------------------------------- #!#
                
                
                
        #*locks the laptop without closing anything
        elif 'lock device' in query:
            print("Locking device")
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()
 
        #*shuts down system completely
        elif 'shutdown system' in query:
            print("Hold On a Sec ! Your system is on its way to shut down")
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')
                 
        #*clears all items in recycle bin
        elif 'empty recycle bin' in query:
            print("Recycling...")
            speak("Recycling")
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            print("Recycle Bin Recycled")
            speak("Recycle Bin Recycled")
 
       #*blocks bot from listening for a defined time asked in seconds
        elif "don't listen" in query or "stop listening" in query:
            print("Alright. How long do you not want me to listen for?")
            speak("Alright. How long do you not want me to listen for?")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
         
        #*restarts laptop   
        elif "restart laptop" in query:
            print("Restarting...")
            speak("Restarting")
            subprocess.call(["shutdown", "/r"])
             
        #*puts computer to sleep
        elif "hibernate" in query or "sleep" in query:
            print("Hibernating...")
            speak("Hibernating...")
            subprocess.call("shutdown / h")

        #*sings out, then logs off
        elif "log off" in query or "sign out" in query:
            print("Make sure all the application are closed before sign-out")
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])
            
        else:
            print("I'm Sorry, I didn't quite get you")
            speak("Im sorry, I didn't quite get you")
