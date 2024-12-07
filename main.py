import math
import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen



engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

assname = ("Nova")



def speak(audio):
	engine.say(audio)
	engine.runAndWait()

def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<12:
		speak("Good Morning!")

	elif hour>= 12 and hour<18:
		speak("Good Afternoon!") 

	else:
		speak("Good Evening!") 

	assname =("Nova")
	speak("I am your Assistant")
	speak(assname)
	

def username():
	# speak("What should i call you sir")
	# uname = takeCommand()
	speak("Welcome Master")
	# speak(uname)
	columns = shutil.get_terminal_size().columns
	
	# print("#####################".center(columns))
	# print("Welcome ", uname.center(columns))
	# print("#####################".center(columns))
	
	speak("How can i Help you?")

def takeCommand():
	
	r = sr.Recognizer()
	
	with sr.Microphone() as source:
		
		print("Listening...")
		r.pause_threshold = 1
		audio = r.listen(source)

	try:
		print("Recognizing...") 
		query = r.recognize_google(audio, language ='en-in')
		print(f"User said: {query}\n")

	except Exception as e:
		print(e) 
		print("Unable to Recognize your voice.") 
		return "None"
	
	return query

def sendEmail(to, content):
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.ehlo()
	server.starttls()
	
	# Enable low security in gmail
	server.login('EnterMail', 'MailKey')
	server.sendmail('EnterMail', to, content)
	server.close()



if __name__ == '__main__':
	clear = lambda: os.system('cls')
	
	# This Function will clean any
	# command before execution of this python file
	clear()
	wishMe()
	username()
	
	while True:
		
		query = takeCommand().lower()
		
		# All the commands said by user will be 
		# stored here in 'query' and will be
		# converted to lower case for easily 
		# recognition of command
		if 'wikipedia' in query:
			speak('Searching Wikipedia...')
			query = query.replace("wikipedia", "")
			results = wikipedia.summary(query, sentences = 3)
			speak("According to Wikipedia")
			print(results)
			speak(results)

		elif 'open youtube' in query:
			speak("Here you go to Youtube\n")
			webbrowser.open("youtube.com")

		elif 'open google' in query:
			speak("Here you go to Google\n")
			webbrowser.open("google.com")

		elif 'open stackoverflow' in query:
			speak("Here you go to Stack Over flow.Happy coding")
			webbrowser.open("stackoverflow.com") 


		elif 'the time' in query:
			strTime = datetime.datetime.now().strftime("%H:%M:%S") 
			speak(f"the time is {strTime}")

		elif 'send a mail' in query or 'send an email' in query:
			try:
				speak("What should I say?")
				content = takeCommand()
				speak("Whom should i send, kindly type email of the receiver")
				print("Email Address : ")
				to = input() 
				sendEmail(to, content)
				speak("Email has been sent !")
			except Exception as e:
				print(e)
				speak("I am not able to send this email")

		elif 'how are you' in query:
			speak("I am fine, Thank you")
			speak("How are you")

		elif 'fine' in query or "good" in query:
			speak("It's good to know that your fine")


		elif "what's your name" in query or "what is your name" in query:
			speak("My friends call me")
			speak(assname)
			print("My friends call me", assname)

		elif 'exit' in query:
			speak("Thanks for giving me your time")
			exit()

		elif "who made you" in query or "who created you" in query: 
			speak("I have been created by Ishan, Ashish and Harry.")
			
		elif 'joke' in query:
			speak(pyjokes.get_joke())
			

		elif "calculate" in query: 
			app_id = "WolframeAPI"
			client = wolframalpha.Client(app_id)
			indx = query.lower().split().index('calculate') 
			query = query.split()[indx + 1:] 
			res = client.query(' '.join(query)) 
			answer = next(res.results).text
			print("The answer is " + answer) 
			speak("The answer is " + answer) 

		elif 'search' in query or 'play' in query:
			query = query.replace("search", "") 
			query = query.replace("play", "")		 
			webbrowser.open(query) 

		elif "who i am" in query:
			speak("If you talk then definitely your human.")

		elif "why you came to world" in query:
			speak("Ishan, Ashish and Harry programmed me.")

	# TODO
		elif "who are you" in query:
			speak("I am your virtual assistant created by Ishan, Ashish and Harry")

		elif 'reason for you' in query:
			speak("I was created as a Minor project by  Ishan, Ashish and Harry")

		elif 'change background' in query:
			ctypes.windll.user32.SystemParametersInfoW(20, 
													0, 
													"BGAddress",
													0)
			speak("Background changed successfully")

		elif 'news' in query:
			try: 
				jsonObj = urlopen('''NewsAPI''')
				data = json.load(jsonObj)
				i = 1
				
				speak('here are some top news from Google News')
				print('''=============== Google News ============'''+ '\n')
				
				for item in data['articles']:
					
					print(str(i) + '. ' + item['title'] + '\n')
					# print(item['description'] + '\n')
					speak(str(i) + '\n' + '. ' + item['title'] + '\n')
					i += 1
					if(i is 6):
						break
			except Exception as e:
				
				print(str(e))


		elif 'shutdown system' in query:
				speak("Hold On a Sec ! Your system is on its way to shut down")
				subprocess.call('shutdown / p /f')
				
		elif 'empty recycle bin' in query:
			winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
			speak("Recycle Bin Recycled")

		elif "don't listen" in query or "stop listening" in query:
			speak("Okay, not hearing for 2 minutes\n")
			a = 120
			time.sleep(a)

		elif "where is" in query:
			query = query.replace("where is", "")
			location = query
			speak("User asked to Locate")
			speak(location)
			webbrowser.open("https://www.google.nl/maps/place/" + location + "")

		elif "restart" in query:
			subprocess.call(["shutdown", "/r"])
			
		elif "hibernate" in query or "sleep" in query:
			speak("Hibernating")
			subprocess.call("shutdown / h")

		elif "write a note" in query:
			speak("What should i write")
			note = takeCommand()
			file = open('note.txt', 'w')
			speak("Should i include date and time")
			snfm = takeCommand()
			if 'yes' in snfm or 'sure' in snfm:
				strTime = datetime.datetime.now().strftime("%H:%M:%S")
				file.write(strTime)
				file.write(" :- ")
				file.write(note)
				speak("The note has been saved successfully")
			else:
				file.write(note)
		
		elif "show note" in query:
			speak("Showing Notes")
			file = open("note.txt", "r") 
			print(file.read())
			speak(file.read(6))
					
		elif "nova" in query or "wish me" in query:
			wishMe()
			speak("Nova 1 in your service Mister")
			speak(assname)

		elif "weather" in query:
			api_key = "OpenWeatherAPI"
			base_url = "http://api.openweathermap.org/data/2.5/weather?"
			speak(" City name ")
			print("City name : ")
			city_name = takeCommand()
			complete_url = base_url + "appid=" + api_key + "&q=" + city_name
			response = requests.get(complete_url) 
			x = response.json() 
			
			if x["cod"] != "404" : 
				y = x["main"] 
				current_temperature = math.trunc(y["temp"] - 273.15)
				current_pressure = y["pressure"] 
				current_humidiy = y["humidity"] 
				feels_like = math.trunc(y["feels_like"] - 273.15)
				z = x["weather"] 
				weather_description = z[0]["description"] 
				speak(" Current Temperature is " +str(current_temperature)+" Degree Celsius and feels like " + str(current_temperature) + "\n humidity is " +str(current_humidiy) +" percent \n Overall it would be " +str(weather_description)) 
				print((" Current Temperature is " +str(current_temperature)+" Degree Celsius and feels like " + str(current_temperature) + "\n humidity is " +str(current_humidiy) +" percent \n Overall it would be " +str(weather_description)))
			
			else: 
				speak(" City Not Found ")

		elif "wikipedia" in query:
			webbrowser.open("wikipedia.com")

		elif "Good Morning" in query:
			speak("A warm" +query)
			speak("How are you Mister")
			speak(assname)

		
		elif "how are you" in query:
			speak("I'm fine, glad you asked me that")

		
		# TODO
		elif "what is" in query or "who is" in query:
			
			try:
				client = wolframalpha.Client("APIkey")
				res = client.query(query)
				answer =next(res.results).text
				print (answer)
				speak (answer)
			except StopIteration:
				print ("No results")

		# elif "" in query:
			# Command go here
			# For adding more commands



