import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import pywhatkit as wk
import time
import threading
import os
import cv2  # OpenCV for face detection
import numpy as np  # For image processing

# Initialize the speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
engine.setProperty('rate', 160)

# Password for authentication
password = "jarvis"  # Set password to "jarvis"

# Global flag to stop the program
stop_flag = False

# Load pre-trained face detection model (Haarcascade)
face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')

def speak(audio):
    print(f"Speaking: {audio}")
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good Morning Boss!")
    elif 12 <= hour < 18:
        speak("Good Afternoon Boss!")
    else:
        speak("Good Evening Boss!")
        
    speak("How are you Boss? What can I do?")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
        
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language="en-IN")
        print(f"User said: {query}\n")
    except Exception as e:
        print(f"Error: {e}")
        print("Say that again, please...")
        speak("Say that again, Boss please...")
        return "None"
    
    return query.lower()

def detect_face():
    # Start video capture
    cap = cv2.VideoCapture(0)
    speak("Looking for a face, please look at the camera.")

    face_detected = False
    for _ in range(100):  # Limit the number of attempts to detect a face
        ret, frame = cap.read()
        if not ret:
            continue
        
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        
        # Detect faces
        faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30))

        # If a face is detected, set the flag to True
        if len(faces) > 0:
            face_detected = True
            for (x, y, w, h) in faces:
                cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)
            speak("Face detected successfully.")
            break

    cap.release()

    if not face_detected:
        speak("No face detected. Exiting program.")
        return False
    return True

def authenticate():
    if detect_face():
        speak("Please enter the password to proceed.")
        password_attempt = takeCommand().strip()
        if password_attempt == password:
            speak("Authentication successful. Starting Jarvis.")
            return True
        else:
            speak("Authentication failed. Exiting program.")
            return False
    else:
        return False

def monitor_stop_command():
    global stop_flag
    while True:
        query = takeCommand()
        if 'stop' in query:
            stop_flag = True
            speak("Stopping all actions.")
            break

if __name__ == "__main__":    
    if not authenticate():
        exit()

    wishMe()
    while True:
        query = takeCommand()

        if query == 'stop':
            break  # Exit the loop and stop the program
        
        elif 'jarvis' in query:
            print("Yes Boss") 
            speak("Yes Boss")
            
        elif 'what is your name' in query:
            print("My name is Jarvis")
            speak("My name is Jarvis")
        
        elif 'what is ' in query:
            speak('Searching Wikipedia...')
            query = query.replace('wikipedia', '')
            try:
                result = wikipedia.summary(query, sentences=5)
                speak('According to Wikipedia')
                print(result)
                speak(result)
            except wikipedia.exceptions.DisambiguationError:
                speak("There are multiple results for this. Please specify more clearly.")
            except wikipedia.exceptions.PageError:
                speak("Sorry, I couldn't find any information on that.")
            except Exception as e:
                speak("An error occurred while searching Wikipedia.")
                print(e)

        elif 'just open google' in query:
            webbrowser.open('https://www.google.com')
            
        elif 'open google' in query:
            speak("What should I search?")
            query = takeCommand()
            webbrowser.open(f'https://www.google.com/search?q={query}')
        
        elif 'just open youtube' in query:
            webbrowser.open('https://www.youtube.com')
        
        elif 'open youtube' in query:
            speak("What should I search?")
            search_query = takeCommand()
            if search_query:
                speak(f"Searching for {search_query} on YouTube...")
                wk.playonyt(search_query)
                
                stop_thread = threading.Thread(target=monitor_stop_command)
                stop_thread.start()
                
                for i in range(200):
                    time.sleep(1)
                    if stop_flag:
                        break
                
                if not stop_flag:
                    speak("The song has ended. What else can I do for you, Boss?")
                else:
                    break

        elif 'just open gmail' in query:
            webbrowser.open('https://mail.google.com')
            
        elif 'open gmail' in query:
            speak("Opening Gmail...")
            webbrowser.open('https://mail.google.com')
            
        elif 'just open instagram' in query:
            speak("Opening Instagram...")
            webbrowser.open('https://instagram.com')
            
        elif 'just open netflix' in query:
            webbrowser.open('https://www.netflix.com')
        
        elif 'open netflix' in query:
            speak("Opening Netflix...")
            webbrowser.open('https://www.netflix.com')
            
        elif 'just open chat gpt' in query:
            webbrowser.open('https://www.chatgpt.com')
            
        elif 'open chat gpt' in query:
            speak("What should I search?")
            query = takeCommand()
            webbrowser.open(f'https://www.chatgpt.com/results?search_query={query}')
            
        elif 'close browser' in query:
            os.system("taskkill /f /im msedge.exe")
            speak("Microsoft Edge has been closed.")
            
        elif 'close chrome' in query:
            os.system("taskkill /f /im chrome.exe")
            speak("Google Chrome has been closed.")
#...............................................................................................................

        
