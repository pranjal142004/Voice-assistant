import os
import pyttsx3
import speech_recognition as sr
import datetime
import pywhatkit as wk
import time
import pyautogui
import pygame
import webbrowser
import wikipedia
import subprocess
import ctypes
import cv2
import sys


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)  
engine.setProperty('rate', 160)

def speak(audio):
    print(f"Speaking: {audio}")
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good Morning Sir!")
    elif 12 <= hour < 18:
        speak("Good Afternoon Sir!")
    else:
        speak("Good Evening Sir!")
    speak("How are you Sir? What can I do?")

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
    except sr.UnknownValueError:
        speak("Sorry, I didn't understand that. Please say it again.")
        return "None"
    except sr.RequestError as e:
        speak("Sorry, the speech recognition service is down.")
        return "None"
    except Exception as e:
        print(f"Error: {e}")
        speak("Say that again, Sir please...")
        return "None"
    return query.lower()

def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = recognizer.listen(source)
        
        try:
            text = recognizer.recognize_google(audio)
            print(f"You said: {text}")
            return text
        except sr.UnknownValueError:
            speak("Sorry, I didn't catch that. Please try again.")
            return ""
        except sr.RequestError:
            speak("There seems to be an issue with the service. Please try again later.")
            return ""

def open_notepad():
    speak("What would you like me to write in Notepad?")
    text = listen()  
    
    speak("Opening Notepad.")
    process = subprocess.Popen('notepad.exe')
    
    time.sleep(1)  
    
    if text:
        speak(f"Writing the text in Notepad: {text}")
        os.system(f'echo {text} | clip')  
        os.system('powershell "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait(\'^v\')"')  # Pastes the text in Notepad
    else:
        speak("No text was provided to write in Notepad.")

def close_notepad():
    os.system("taskkill /f /im notepad.exe")
    speak("Notepad has been closed.")

def open_commandprompt():
    speak("Opening Command Prompt.")
    try:
        subprocess.Popen('start cmd', shell=True)
    except Exception as e:
        speak("Failed to open Command Prompt.")
        print(f"Error: {e}")

def close_commandprompt():
    import os
    os.system("taskkill /f /im cmd.exe")
    speak("Command Prompt has been closed.")

def open_paint():
    speak("Opening Paint.")
    subprocess.Popen('mspaint.exe')

def close_paint():
    os.system("taskkill /f /im mspaint.exe")
    speak("Paint has been closed.")

def open_word():
    word_path = "C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE"  
    if os.path.exists(word_path):
        speak("Opening Microsoft Word.")
        subprocess.Popen([word_path])
    else:
        speak("Microsoft Word is not installed or the path is incorrect.")

def close_word():
    os.system("taskkill /f /im WINWORD.EXE")
    speak("Microsoft Word has been closed.")

def open_excel():
    excel_path = "C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE"  
    if os.path.exists(excel_path):
        speak("Opening Microsoft Excel.")
        subprocess.Popen([excel_path])
    else:
        speak("Microsoft Excel is not installed or the path is incorrect.")

def close_excel():
    os.system("taskkill /f /im EXCEL.EXE")
    speak("Microsoft Excel has been closed.")

def open_calculator():
    speak("Opening Calculator.")
    subprocess.Popen('calc.exe')

def close_calculator():
    os.system("taskkill /f /im Calculator.exe")
    speak("Calculator has been closed.")

def take_screenshot():
    speak("Taking a screenshot...")
    
    screenshot_directory = "screenshots"
    
    if not os.path.exists(screenshot_directory):
        os.makedirs(screenshot_directory)
    
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    screenshot_path = os.path.join(screenshot_directory, f"screenshot_{timestamp}.png")
    
    screenshot = pyautogui.screenshot()
    screenshot.save(screenshot_path)
    
    speak(f"Screenshot taken and saved as {screenshot_path}.")
    
    speak("Opening the gallery for you.")
    subprocess.Popen(f'explorer "{os.path.abspath(screenshot_directory)}"')

def open_gallery():
    speak("Opening your picture gallery.")
    pictures_folder = os.path.join(os.environ["USERPROFILE"], "Pictures")  
    subprocess.Popen(f'explorer "{pictures_folder}"')

def close_gallery():
    os.system("taskkill /f /im explorer.exe")
    speak("Gallery has been closed.")

def shutdown_system():
    speak("Shutting down the system.")
    os.system("shutdown /s /t 1")

def restart_system():
    speak("Restarting the system.")
    os.system("shutdown /r /t 1")

def lock_system():
    speak("Locking the system.")
    ctypes.windll.user32.LockWorkStation()

def calculate(expression):
    try:
        result = eval(expression)
        speak(f"The result is {result}")
    except Exception as e:
        speak("Sorry, I couldn't perform that calculation.")
        print(e)

if __name__ == "__main__":    
    wishMe()
    while True:
        query = takeCommand()
        if query == 'stop':
            break  
        elif 'jarvis' in query:
            speak("Yes Boss")
        
        elif 'what is your name' in query:
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
            chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"
            url = "https://www.google.com"
            if os.path.exists(chrome_path):
                subprocess.Popen([chrome_path, url])
            else:
                speak("I couldn't find Google Chrome. Please check the installation.")

        elif 'open google' in query:
            speak("What should I search?")
            search_query = takeCommand()
            if search_query:
                webbrowser.open(f'https://www.google.com/search?q={search_query}')
                
        elif 'just open youtube' in query:
            chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"
            url = "https://www.youtube.com"
            if os.path.exists(chrome_path):
                subprocess.Popen([chrome_path, url])
            else:
                speak("I couldn't find Google Chrome. Please check the installation.")
        
        elif 'open youtube' in query:
            speak("What should I search?")
            search_query = takeCommand()
            if search_query:
                speak(f"Searching for {search_query} on YouTube...")
                wk.playonyt(search_query)
                speak("Enjoy the video, Boss!")
                
                time.sleep(180)
                speak("It's been 3 minutes, Sir. Do you need anything else?")
                next_query = takeCommand()

        elif 'open gmail' in query:
            speak("Opening Gmail...")
            webbrowser.open('https://mail.google.com')
            
        elif 'open instagram' in query:
            speak("Opening Instagram...")
            webbrowser.open('https://instagram.com')
            
        elif 'open netflix' in query:
            speak("Opening Netflix...")
            webbrowser.open('https://www.netflix.com')
            
        elif 'close browser' in query:
            os.system("taskkill /f /im msedge.exe")
            speak("Microsoft Edge has been closed.")
            
        elif 'close chrome' in query:
            os.system("taskkill /f /im chrome.exe")
            speak("Google Chrome has been closed.")
            
        elif 'open notepad' in query:
            open_notepad()
        elif 'close notepad' in query:
            close_notepad()
        elif 'open command prompt' in query:
            open_commandprompt()
        elif 'close command prompt' in query:
            close_commandprompt()
        elif 'open paint' in query:
            open_paint()
        elif 'close paint' in query:
            close_paint()
        elif 'open word' in query:
            open_word()
        elif 'close word' in query:
            close_word()
        elif 'open excel' in query:
            open_excel()
        elif 'close excel' in query:
            close_excel()
        elif 'open calculator' in query:
            open_calculator()
        elif 'close calculator' in query:
            close_calculator()
        elif 'take screenshot' in query:
            take_screenshot()
        elif 'shutdown the system' in query:
            shutdown_system()
        elif 'restart the system' in query:
            restart_system()
        elif 'lock the system' in query:
            lock_system()
        elif 'open gallery' in query:
            open_gallery()
        elif 'close gallery' in query:
            close_gallery()
        elif 'what is the time' in query:  
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")
            
        elif 'volume up' in query:
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            pyautogui.press("volumeup")
            
        elif 'volume down' in query:
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            pyautogui.press("volumedown")
            
        elif 'open camera' in query:
            cap = cv2.VideoCapture(0)
            while True:
                ret, img = cap.read()
                cv2.imshow('webcam', img)
                k = cv2.waitKey(50)
                if k==27:
                    break;
                cap.release()
                cv2.destroyAllWindows()
                
        elif 'go to sleep' in query:
            speak(' alright then, I am switiching off')
            sys.exit()
                
            
