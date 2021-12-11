
#!--------------------------Module Import--------------------------!#

try:
    #The subprocess module allows spawning of new processes, connecting to their input/output/error pipes, and obtaining their return codes.
    import subprocess
    #pyttsx3 is a python text to speech library.
    import pyttsx3
    #json module is a javascrpit object decoder for python
    import json
    #random module for giving
    # random results
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
    #Python web scraper
    from bs4 import BeautifulSoup
    #module for windows application automation
    import win32com.client as wincl
    #module to open links
    from urllib.request import urlopen
    #turtle module for gui related commands
    from turtle import *
    #freegames module for simple games
    from freegames import line
    #Coingecko API wrapper for python to get cryptocurrency prices
    import pycoingecko
    #Coingecko API https get requester
    from pycoingecko import CoinGeckoAPI
    cg = CoinGeckoAPI()
    import settings
except:
    print("Import error. Please open a github issue")

#!--------------------------User Variables--------------------------!#
try:
    browser = settings.browser
    username = settings.username
    password = settings.password
    currency = settings.currency
    voice_setting = settings.voice_setting
    assistantname = settings.assistantname


    engine = settings.engine
    voices = settings.voices
    #sets output voice to male or female.
    engine.setProperty("voice", voices[voice_setting].id)
except:
     print("Variable Error. Please open a github issue")
     
#!--------------------------Base Function--------------------------!#
     
#function for voice output 
def speak(audio=""):
    engine.say(audio)
    engine.runAndWait()

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

def BugReport():
    print("Command Failed. Please open a github issue")
    speak("Command Failed. Please open a github issue")
    webbrowser.get(browser).open("https://github.com/RandomPerson06/VirtualAssistant.py/blob/main/VoiceAssistant.py")


#clears previous text from terminal and executes the Greeting function
if __name__ == '__main__':
    clear = lambda: os.system('cls')

    clear()
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        print("Good Morning")
        speak("Good Morning")
    elif hour >= 12 and hour < 18:
        print("Good Afternoon")
        speak("Good Afternoon")
    elif hour >= 18 and hour < 20:
        print("Good Evening")
        speak("Good Evening")
    else:
        print("Hello")
        speak("Hello")
    print("I am Your Assistant, " + assistantname)
    speak("I Am Your Assistant, " + assistantname)
 
        
#once query is registered it is registered and processed here to give an output
while 0==0:
        
        #query is moved to all lowercase for easier recognition by the API
        query = takeCommand().lower()
        
  #!#--------------------------------------------Link Opening Commands-------------------------------------------- #!#        

        #open google.com
        if query == 'open google':
            try:
                print("Opening Google...\n")
                speak("Opening Google.com\n")
                webbrowser.get(browser).open("google.com")
            except:
                BugReport()
    
        #search google
        elif "search" in query and "for" in query and "google" in query:
            try:
                query = query.replace("search google for", "")
                query = query.replace("search for", "")
                query = query.replace("search", "")
                query = query.replace("on chrome", "")
                query = query.replace("on google", "")
                query = query.replace("google", "")

                print("Googling " + query)
                speak("Googling " + query)
                webbrowser.get(browser).open("https://www.google.com/search?q=" + query)
            except:
                BugReport()
                
        #search wikipedia for query
        elif 'wikipedia' in query:
            try:
                print("Searching Wikipedia...")
                speak('Searching Wikipedia...')
                query = query.replace("what does", "")
                query = query.replace("search wikipedia for", "")
                query = query.replace("search wikipedia", "")
                query = query.replace("search for", "")
                query = query.replace("on wikipedia", "")
                query = query.replace("wikipedia", "")
                results = wikipedia.summary(query, sentences = 3)
                print("According to Wikipedia...\n")
                speak("According to Wikipedia")
                print(results)
                speak(results)
            except:
                BugReport()
                
        #open youtube.com
        elif query == 'open youtube':
            try:
                print("Opening Youtube...\n")
                speak("Opening Youtube...\n")
                webbrowser.get(browser).open("youtube.com")
            except:
                BugReport()

        #search and play on youtube
        elif ("play" in query or ("search" in query)) and "on youtube" in query:
            try:
                query = query.replace ("play", "")
                query = query.replace ("on youtube", "")
                print("Searching for " + query + "on Youtube")
                speak("Searching for " + query + "on Youtube")
                search = query
                webbrowser.get(browser).open("https://www.youtube.com/results?search_query=" + search)
            except:
                BugReport()                
            
        #search and play on spotify    
        elif ("play" in query or "search" in query) and "on spotify" in query:
            try:
                query = query.replace("play", "")
                query = query.replace("on spotify", "")
                print("Searching Spotify for " + query)
                speak("Searching spotify for " + query)
                search = query
                webbrowser.get(browser).open("https://open.spotify.com/search/" + search)
            except:
                BugReport()
                
        #search and play on spotify        
        elif ("play" in query or "search" in query) and ("on gaana" or "on gana" or "on ganna") in query:
            try:
                query = query.replace("play", "")
                query = query.replace("on", "")
                query = query.replace("gaana", "")
                query = query.replace("gana", "")
                query = query.replace("ganna", "")
                print("Searching Gaana for: " + query)
                speak("Searching Gaana for: " + query)
                search = query
                webbrowser.get(browser).open("https://gaana.com/search/" + search)
            except:
                BugReport()
                                
        #search on google maps
        elif "where is" in query:
            try:
                query = query.replace("where is", "")
                print("Searching for " + query + " on Google Maps")
                speak("Searching for " + query + " on Google Maps")
                location = query
                webbrowser.open("https://www.google.com/maps/place/" + location)
            except:
                BugReport()
                    
        #open binance.com
        elif "open" in query and "binance" in query:
            try:
                print("Opening Binance")
                speak("Opening Binance")
                webbrowser.get(browser).open("https://www.binance.com")
            except:
                BugReport()
        
        #open github.com
        elif "open" in query and "github" in query:
            try:
                print("Opening Github")
                speak("Opening Github")
                webbrowser.get(browser).open("https://www.github.com")
            except:
                BugReport()
            
        #open instagram.com
        elif "open" in query and "instagram" in query:
            try:
                print("Opening Instagram")
                speak("Opening Instagram")
                webbrowser.get(browser).open("https://www.instagram.com")
            except:
                BugReport()
            
        #open twitter.com
        elif "open" in query and "twitter" in query:
            try:
                print("Opening Twitter")
                speak("Opening Twitter")
                webbrowser.get(browser).open("https://www.twitter.com")
            except:
                BugReport()
        
        #open facebook.com
        elif "open" in query and "facebook" in query:
            try:
                print("Opening")
                speak("Opening")
                webbrowser.get(browser).open("https://www.facebook.com")
            except:
                BugReport()
            
        #open discord.com
        elif "open" in query and "discord" in query:
            try:
                print("Opening Discord")
                speak("Opening Discord")
                webbrowser.get(browser).open("https://discord.com/channels/@me")
            except:
                BugReport()
            
        #open whatsapp.com
        elif "open" in query and "whatsapp" in query:
            try:
                print("Opening Whatsapp")
                speak("Opening Whatsapp")
                webbrowser.get(browser).open("https://web.whatsapp.com")
            except:
                BugReport()
        
        #open gmail.com
        elif "open" in query and ("mail" or "gmail") in query:
            try:
                print("Opening Mail")
                speak("Opening Mail")
                webbrowser.get(browser).open("https://gmail.com")
            except:
                BugReport()
                        
    
 #!#--------------------------------------------High Intensity Commands-------------------------------------------- #!#


        elif "connect4" in query or "connect 4" in query or "connect four" in query:
            try:
                print("Opening Connect 4")
                speak("Opening Connect Four")
                speak("Connect 4 is now open. If you can't see it, please check for a new window")
                def connect4():
                    #The connect 4 command was not made by me. It is one of the default open source games from the freegames python library
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
                    
                connect4()
            except:
                BugReport()

                
        elif "send" in query and "mail" in query:
            try:
                if (username != "") and (password != ""):
                    print("Voice recognition is not supported for mail yet. Please check the terminal for sending a mail")
                    speak("Voice recognition is not supported for mail yet. Please check the terminal for sending a mail")
                    def sendEmail(to, content):
                        server = smtplib.SMTP('smtp.gmail.com', 587)
                        server.ehlo()
                        server.starttls()
         
                        server.login(username, password)
                        server.sendmail(username, to, content)
                        server.close()
                    print("Enter Recipient's address")
                    speak("Enter Recipient's address")
                    to = input("Recipient's Address: ")
                    print("Enter the message")
                    speak("Enter the message")
                    content = input("Content: ")
                    try:
                        sendEmail(to, content)
                        print("Mail Sent!")
                        speak("Mail Sent!")
                    except:
                        print("Error sending email. Please go to https://www.google.com/settings/security/lesssecureapps and enable low security apps")
                        speak("Error sending email. Please go to https://www.google.com/settings/security/lesssecureapps and enable low security apps")
                else:
                    print("It seems like you have not logged in. Please enter your username and password below")
                    speak("It seems like you have not logged in. Please enter your username and password below")
                    new_username = input("Username: ")
                    new_password = input("Password: ")
                    with open(r"C:\Users\buzzp\Documents\PythonCodes\Miscellaneous\VirtualAssistant.py\settings.py", "w") as file1:
                                file1.write("import pyttsx3 \n")
                                file1.write("chrome = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'\n")
                                file1.write("edge = 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe %s'\n")
                                file1.write("brave = 'C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe %s'\n")
                                file1.write("engine = pyttsx3.init('sapi5')\n")
                                file1.write("voices = engine.getProperty('voices')\n")
                                file1.write("\n")
                                file1.write("browser = " + "'" + settings.browser + "'" + " \n")
                                file1.write("username = " + "'" + new_username + "'" + " \n")
                                file1.write("password = " + "'" + new_password + "'" + "\n")
                                file1.write("currency = " + "'" + settings.currency + "'" + "\n")
                                file1.write("voice_setting = " + settings.voice_setting + "\n")
                                file1.write("assistantname = " + "'" + settings.assistantname + "'" + "\n")
                    file1.close()
            except:
                BugReport()

                
 #!#--------------------------------------------General Questions-------------------------------------------- #!#
        
                #if 'the time' is heard in query, it will display time using datetime module    
        elif 'the time' in query:
            try:
                now = datetime.now()
                strTime = now.strftime("%H:%M.%S")
                print(f"The time is {strTime}")
                Time = now.strftime("%H:%M")
                speak(f"The time is {Time}")
            except:
                BugReport()
                
        #if 'how are you' is heard in query, it says taht it is fine
        elif 'how are you' in query or "how are you" in query or "hru" in query or "how r u" in query:
            try:
                print("I am fine, Thank You")
                speak("I am fine, Thank You")
            except:
                BugReport()
        
        #if anything is heard about it's name, it uses the assistantname var defined earlier
        elif "what's your name" in query or "what is your name" in query:
            try:
                print("My friends call me", assistantname)
                speak("My friends call me ")
                speak(assistantname)
            except:
                BugReport()     

        #search on google
        elif "search for" in query or "what is" in query:
            try:
                query = query.replace("search for", "")
                print ("Searching for: " + query)
                speak("Searching for: " + query)
                search = query
                webbrowser.get(browser).open("https://www.google.com/search?q=" + query) 
            except:
                BugReport()
       

 #!#-------------------------------------------Cryptocurrency Commands-------------------------------------------- #!#
            
        

      
        elif "bitcoin price" in query or "price of bitcoin" in query or "btc price" in query or "price of btc" in query:
            try:
                cryptoprice = cg.get_price(ids="bitcoin", vs_currencies=currency)
                price = cryptoprice["bitcoin"]  
                print("Bitcoin"+ " is currently worth " + str(price[currency]) + " " + currency)
                speak("Bitcoin" + " is currently worth " + str(price[currency]) + " " + currency)
            except:
                BugReport()

        elif "ethereum price" in query or "price of ethereum" in query or "eth price" in query or "price of eth" in query:
            try:
                cryptoprice = cg.get_price(ids="ethereum", vs_currencies=currency)
                price = cryptoprice["ethereum"]  
                print("Ethereum"+ " is currently worth " + str(price[currency]) + " " + currency)
                speak("Ethereum" + " is currently worth " + str(price[currency]) + " " + currency)
            except:
                BugReport()
           
        elif "price" in query:
            print("What crypto currency would you like to find the price of?")
            speak("What crypto currency would you like to find the price of?")

            
            with sr.Microphone() as source:
                sr.Recognizer().pause_threshold = 1
                audio = sr.Recognizer().listen(source)
            try:
                query = sr.Recognizer().recognize_google(audio, language ='en-in')
                query = query.lower()
            except Exception as e:
                print("Unable to Recognize your voice.")
                
            print("Searching for price of " + query)
            speak("Searching for price of " + query)

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
        
        
         #if joke is heard, it gives a python joke
        elif "joke" in query or "riddle" in query:
            try:
                joke = pyjokes.get_joke("en","all")
                print(joke)
                speak(joke)
            except:
                BugReport()
               
 #!#--------------------------------------------Device Control-------------------------------------------- #!#
                
                
                
        #locks the laptop without closing anything
        elif 'lock device' in query:
            try:
                print("Locking device")
                speak("locking device")
                ctypes.windll.user32.LockWorkStation()
            except:
                BugReport()
 
        #shuts down system completely
        elif 'shutdown system' in query:
            try:
                print("Are you sure you want to shut down?")
                speak("Are you sure you want to shut down?")
                with sr.Microphone() as source:
                        print("Listening...")
                        sr.Recognizer().pause_threshold = 1
                        audio = sr.Recognizer().listen(source)
                try:
                        query = sr.Recognizer().recognize_google(audio, language ='en-in')
                        query = query.lower()
                except Exception as e:
                    print("Unable to Recognize your voice.")
                if query == ("y" or "yes" or "sure" or "k" or "ok"):
                    print("Hold On a Sec ! Your system is on its way to shut down")
                    speak("Hold On a Sec ! Your system is on its way to shut down")
                    subprocess.call('shutdown / p /f')
                else:
                    print("Cancelled")
                    speak("Cancelled")
            except:
                BugReport()    
                 
        #clears all items in recycle bin
        elif 'recycle bin' in query:
            print("Recycling...")
            speak("Recycling")
            winshell.recycle_bin().empty(confirm = True, show_progress = True, sound = True)
            print("Recycle Bin Recycled")
            speak("Recycle Bin Recycled")
 
        #restarts laptop   
        elif "restart laptop" in query:
            try:
                print("Are you sure you want to restart?")
                speak("Are you sure you want to restart?")
                with sr.Microphone() as source:
                        print("Listening...")
                        sr.Recognizer().pause_threshold = 1
                        audio = sr.Recognizer().listen(source)
                try:
                        query = sr.Recognizer().recognize_google(audio, language ='en-in')
                        query = query.lower()
                except Exception as e:
                    print("Unable to Recognize your voice.")
                if query == ("y" or "yes" or "sure" or "k" or "ok"):
                    print("Restarting...")
                    speak("Restarting")
                    subprocess.call(["shutdown", "/r"])
                else:
                    print("Cancelled")
                    speak("Cancelled")
            except:
                BugReport()
                    
       
        elif "hibernate" in query or "sleep" in query:
            try:
                print("Hibernating...")
                speak("Hibernating...")
                subprocess.call("shutdown /h")
            except:
                BugReport()
                
        #sings out, then logs off
        elif "log off" in query or "sign out" in query:
            try:
                print("Are you sure you want to Sign out?")
                speak("Are you sure you want to Sign out?")
                with sr.Microphone() as source:
                        print("Listening...")
                        sr.Recognizer().pause_threshold = 1
                        audio = sr.Recognizer().listen(source)
                try:
                        query = sr.Recognizer().recognize_google(audio, language ='en-in')
                        query = query.lower()
                except Exception as e:
                    print("Unable to Recognize your voice.")
                if query == ("y" or "yes" or "sure" or "k" or "ok"):
                    print("Logging out")
                    speak("Logging out")
                    subprocess.call(["shutdown", "/l"])
                else:
                    print("Cancelled")
                    speak("Cancelled")
            except:
                BugReport()    
            
 #!#--------------------------------------------Assistant Settings-------------------------------------------- #!#


        elif query == "settings" or query == "assistant settings" or query == "setting":
            try:
                print("What settings would you like to change? You can also input 'help' to see all the settings or 'cancel' to go back to the assistant")
                speak("What settings would you like to change?")
                with sr.Microphone() as source:
                        print("Listening...")
                        sr.Recognizer().pause_threshold = 1
                        audio = sr.Recognizer().listen(source)
                try:
                        query = sr.Recognizer().recognize_google(audio, language ='en-in')
                        query = query.lower()
                        print(f"User said {query}")
                except Exception as e:
                    print("Unable to Recognize your voice.")
                
                if query == 'help' or query == "health":
                    print("You can change the settings of my voice, default currency, default browser and my name. The settings will be printed into the output terminal shortly")
                    speak("You can change the settings of my voice, default currency, default browser and my name. The settings will be printed into the output terminal shortly")
                    
                    print("Voice Settings: input 'voice settings' or 'voice'. More details will be given upon initiating the setting. \n")
                    print("Browser Settings: input 'browser settings' or 'browser'. More details will be given upon initiating the setting \n")
                    print("Currency Settings: input 'currency settings' or 'currency'. More details will be given upon initiating the setting \n")
                    print("Name Settings: input 'name settings' or 'name'. More details will be given upon initiating the setting \n")
                    
                    
                elif query == "voice settings" or query == "voice":
                    print("Sorry, but my voice settings can only be changed directly from the source code. Would you like to know how to change it?")
                    speak("Sorry, but my voice settings can only be changed directly from the source code. Would you like to know how to change it?")
                    with sr.Microphone() as source:
                        print("Listening...")
                        sr.Recognizer().pause_threshold = 1
                        audio = sr.Recognizer().listen(source)
                    try:
                        query = sr.Recognizer().recognize_google(audio, language ='en-in')
                        query = query.lower()
                    except Exception as e:
                        print("Unable to Recognize your voice.")                
                    if query == ("y" or "yes" or "sure" or "k" or "ok"):
                        print("Ok. Go to line 58, and change the voice_setting to 0 for a male voice and 1 for a female voice. More voices will come in the future")
                        speak("Ok. Go to line 58, and change the voice_setting to 0 for a male voice and 1 for a female voice. More voices will come in the future")
                    else:
                        print("Ok. Returning back to assistant mode")
                        speak("Ok. Returning back to assistant mode")
                
                elif query == "currency settings" or query == "currency" or query == "currency setting":
                    print("Currency can be changed to any standard 3 letter currency abbreviation")
                    speak("Currency can be changed to any standard 3 letter currency abbreviation")

                    with sr.Microphone() as source:
                        print("Listening...")
                        sr.Recognizer().pause_threshold = 1
                        audio = sr.Recognizer().listen(source)
                    try:
                        query = sr.Recognizer().recognize_google(audio, language ='en-in')
                        query = query.lower()
                    except Exception as e:
                        print("Unable to Recognize your voice.")

                    print("Unable to Recognize your voice.")
                    currency = query
                    print(currency)
                    print("Attempted to change currency. Please do the crypto currency command to check if it is a valid code")
                    speak("Attempted to change currency. Please do the crypto currency command to check if it is a valid code")
                        
                elif query == "browser setting" or query == "browser" or query == "browser settings":
                    print("Sorry, but my browser settings can only be changed directly from the source code. Would you like to know how to change it?")
                    speak("Sorry, but my browser settings can only be changed directly from the source code. Would you like to know how to change it?")
                    with sr.Microphone() as source:
                        print("Listening...")
                        sr.Recognizer().pause_threshold = 1
                        audio = sr.Recognizer().listen(source)
                    try:
                        query = sr.Recognizer().recognize_google(audio, language ='en-in')
                        query = query.lower()
                    except Exception as e:
                        print("Unable to Recognize your voice.")
                    if query == ("y" or "yes" or "sure" or "k" or "ok"):
                        print("Ok. Go to line 58, and change the voice_setting to 0 for a male voice and 1 for a female voice. More voices will come in the future")
                        speak("Ok. Go to line 58, and change the voice_setting to 0 for a male voice and 1 for a female voice. More voices will come in the future")
                    else:
                        print("Ok. Returning back to assistant mode")
                        speak("Ok. Returning back to assistant mode")            

                    browser = query
                    print("Successfully Attempted to change browser")
                    speak("Successfully Attempted to change browser")
                    
                elif query == "name setting" or query == "name settings" or query == "name" or query == "neem":
                    print("My name can be set to anything you want!")
                    speak("My name can be set to anything you want!")

                    with sr.Microphone() as source:
                        print("Listening...")
                        sr.Recognizer().pause_threshold = 1
                        audio = sr.Recognizer().listen(source)
                    try:
                        query = sr.Recognizer().recognize_google(audio, language ='en-in')
                        query = query.lower()
                        print(f"User said {query}")
                    except Exception as e:
                        print("Unable to Recognize your voice.")


                    assistantname = query
                    print("Succesfully set new name!")
                    speak("Succesfully set new name!")
                    
                elif query == 'cancel':
                    speak("Ok, returning back to Assistant")
                    print("Ok, returning back to Assistant")
                    
                else:
                    print("I'm sorry. I couldn't understand that. Returning back to Assistant mode")
                    speak("Im sorry. I couldnt understand that. Returning back to Assistant mode")
            except:
                BugReport()

 #!#--------------------------------------------End Point-------------------------------------------- #!#

 
        else:
            print("I'm Sorry, I didn't quite get you. Would you like to add this feature?")
            #speak("I'm Sorry, I didn't quite get you. Would you like to add this feature?")
            with sr.Microphone() as source:
                    print("Listening...")
                    sr.Recognizer().pause_threshold = 1
                    audio = sr.Recognizer().listen(source)
            try:
                     query = sr.Recognizer().recognize_google(audio, language ='en-in')
                     query = query.lower()
            except Exception as e:
                    print("Unable to Recognize your voice.")
            if query == ("y" or "yes" or "sure" or "k" or "ok"):
                print("Ok. Please open a issue on github and the feature will be added as soon as possible")
                speak("Ok. Please open a issue on github and the feature will be added as soon as possible")
                webbrowser.get(browser).open("https://github.com/RandomPerson06/VirtualAssistant.py/issues/new")
            else:
                print("Ok. Returning back to assistant mode")
                speak("Ok. Returning back to assistant mode")