import pyttsx3 
chrome = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
edge = 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe %s'
brave = 'C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe %s'
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')

browser = chrome
username = '' 
password = ''
currency = 'usd'
voice_setting = 1
assistantname = 'Google'
audio = True
assistant_online = True