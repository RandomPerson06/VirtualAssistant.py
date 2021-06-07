#The subprocess module allows spawning of new processes, connecting to their input/output/error pipes, and obtaining their return codes.
import subprocess
#WolframAlpha is an API which can compute answers using Wolfram’s knowledgebase. It is made possible by the Wolfram Language.
import wolframalpha
#pyttsx3 is a python text to speech library.
import pyttsx3
#tkinter is a gui creator for python
import tkinter
#json module is a javascrpit object decoder for python
import json
#random module for gicing random results
import random
#Python Arthimetic operators for math
import operator
#Python Speech Recognition library
import speech_recognition as sr
#datetime module to give the date and time
import datetime
#wikipedia module to give defenitions
import wikipedia
#webbrowser module to open links
import webbrowser
import os
import winshell
#python jokes
import pyjokes
#module to bring rss feeds for news
import feedparser
#smpt module to send mails
import smtplib
#binds c language data types
import ctypes
#time module to give time
import time
#https requests module to receive data from websites
import requests
#module to edit folder data
import shutil
#twilio module to connect to twilio API, used to receive and send sms, voice calls, etc
from twilio.rest import Client
#Python web scraper
from bs4 import BeautifulSoup
#module for windows application automation
import win32com.client as wincl
#module to open links
from urllib.request import urlopen

#browser settings for opening links
chrome = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
edge = 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe %s'


#change edge to your main browser
browser = edge

#Sets tts settings such as output voice and input type
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
#sets output voice to male or female. voices[1] is for female, whereas it can be changed to voices[0] for male
engine.setProperty("voice", voices[1].id)
#sets a name for the assistant
assistantname = ("google")

#function for voice input 
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

#Says greetings upon bot startup
def wishMe():
    #takes the current hour using datetime module
    hour = int(datetime.datetime.now().hour)
    
    #if the received time is from 12:01AM to 11:59AM it says and prints "Good Morning"
    if hour >= 0 and hour < 12:
        speak("Good Morning")
        print("Good Morning")
        
    #if the received time is from 12:00pm to 5:59pm it says and prints "Good Afternoon"
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon")
        print("Good Afternoon")
        
    #if the received time is from 6:00pm to 7:59pm it says and prints "Good Evening"
    elif hour >= 18 and hour < 20:
        speak("Good Evening")
       
    #if it received any of the times not mentioned above, it just prints "Hello"  
    else:
        speak("Hello")
    
    #Says and prints it's name
    speak("I Am Your Assistant")
    speak(assistantname)

#the function used to take and process voice command
def takeCommand():
     
     #uses sr module to check if there is any input  
    r = sr.Recognizer()
     
    #takes device microphone as input
    with sr.Microphone() as source:
        
        #print's "Listening..." until it receives audio, and wait's 1 second after audio
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
  
    try:
        #uses google speech recognition API to register query
        print("Recognizing...")   
        query = r.recognize_google(audio, language ='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        #if no voice is heard or is not able to understand, it prints "Unable to Recognize your voice" 
        print(e)   
        print("Unable to Recognize your voice.") 
        return "None"
     
    return query

#The sendMail function below is still underprogress
'''
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
     
    # Enable low security in gmail
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()
'''

#clears previous text from terminal and executes the Greeting function
if __name__ == '__main__':
    clear = lambda: os.system('cls')

    clear()
    wishMe()
    
#once query is registered it is registered and processed here to give an output
while True:
        
        #query is moved to all lowercase for easier recognition by the API
        query = takeCommand().lower()
         
        #if wikipedia is heard in the query, it will search wikipedia for the term given 
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences = 3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        #if query strictly says to open youtube, it is done
        elif query == 'open youtube':
            speak("Opening Youtube...\n")
            webbrowser.get('chrome').open("youtube.com")

        #if query strictly says to open google, it is done
        elif query == 'open google':
            speak("Opening Google.com\n")
            webbrowser.open("google.com")
        
        #if 'the time' is heard in query, it will display time using datetime module    
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("% H:% M:% S")   
            speak(f"The time is {strTime}")
            
        #if 'how are you' is heard in query, it says taht it is fine
        elif 'how are you' in query:
            speak("I am fine, Thank you")
        
        #if anything is heard about it's name, it uses the assistantname var defined earlier
        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak(assistantname)
            print("My friends call me", assistantname)
            
        #if joke is heard, it gives a python joke
        elif 'joke' in query:
            speak(pyjokes.get_joke())
        
        #uses wolframapi to do calculations 
        elif "calculate" in query:
             
            app_id = "Wolframalpha api id"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)
            
        #uses browser settings defined arlier to search google
        elif 'search' in query or 'play' in query:
             
            query = query.replace("search", "")
            query = query.replace("play", "")         
            webbrowser.open(query)
            
        #uses json to get rss feed from Times of India
        elif 'news' in query:
             
            try:
                jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                data = json.load(jsonObj)
                i = 1
                 
                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============'''+ '\n')
                 
                for item in data['articles']:
                     
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                 
                print(str(e))
                
        #locks the laptop without closing anything
        elif 'lock device' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
 
        #shuts down system completely
        elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
                 
        #clears all items in recycle bin
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")
 
       #blocks bot from listening for a defined time asked in seconds
        elif "don't listen" in query or "stop listening" in query:
            speak("Alright. How long do you not want me to listen for?")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
            
        #uses webbrowser modules to open a maps.google.com site
        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")
         
        #restarts laptop   
        elif "restart laptop" in query:
            subprocess.call(["shutdown", "/r"])
             
        #puts computer to sleep
        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")
 
        #sings out, then logs off
        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])
            
        #takes weather for a place using openweathermap API
        elif "weather" in query:
             
            # Google Open weather website
            # to get API of Open weather
            api_key = "Api key"
            base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
            speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()
             
            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
             
            else:
                speak(" City Not Found ")
