import speech_recognition as sr
import pyttsx3
import platform
import subprocess
import psutil
import os
import webbrowser
from nltk.corpus import wordnet
import nltk

# Download WordNet dataset once
nltk.download('wordnet')

# Initialize recognizer and TTS engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()

# Set faster voice rate
engine.setProperty('rate', 175)

# Track opened apps
opened_apps = {}

# Adjust for ambient noise once at startup
with sr.Microphone() as source:
    recognizer.adjust_for_ambient_noise(source, duration=1)

def speak(text):
    print(f">> {text}")
    engine.say(text)
    engine.runAndWait()

def listen():
    with sr.Microphone() as source:
        print("Listening...")
        try:
            audio = recognizer.listen(source, timeout=3, phrase_time_limit=5)
            command = recognizer.recognize_google(audio)
            print(f"You said: {command}")
            return command.lower()
        except sr.WaitTimeoutError:
            speak("No input detected.")
        except sr.UnknownValueError:
            speak("Sorry, I didn't catch that.")
        except sr.RequestError:
            speak("Internet connection issue.")
        return ""

# ---------------------- APP LAUNCHING -----------------------

def open_app(app_name):
    system = platform.system()
    if system == "Windows":
        open_windows_app(app_name)
    elif system == "Darwin":
        open_mac_app(app_name)
    elif system == "Linux":
        open_linux_app(app_name)
    else:
        speak(f"Unsupported system: {system}")

def open_windows_app(app_name):
    app_name = app_name.lower()
    common_apps = {
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
        "word": "winword.exe",
        "excel": "excel.exe",
        "chrome": "chrome.exe",
        "firefox": "firefox.exe",
        "edge": "msedge.exe",
        "vlc": "vlc.exe",
        "paint": "mspaint.exe",
        "file explorer": "explorer.exe",
        "vs code": "Code.exe",
        "settings": "ms-settings:",
        "copilot": "ms-copilot:",
        "photos": "microsoft.windows.photos:",
        "onedrive": "OneDrive.exe",
        "microsoft store": "ms-windows-store:",
    }

    if app_name in common_apps:
        app_exe = common_apps[app_name]
        try:
            if app_exe.endswith(":"):
                os.startfile(app_exe)
            elif app_exe.endswith(".exe"):
                subprocess.Popen(app_exe)
            else:
                webbrowser.open(app_exe)
            opened_apps[app_name] = None
            speak(f"Opened {app_name}")
        except Exception as e:
            speak(f"Couldn't open {app_name}. Error: {e}")
    else:
        try:
            subprocess.Popen([app_name])
            opened_apps[app_name] = None
            speak(f"Opened {app_name}")
        except FileNotFoundError:
            speak(f"{app_name} not found. Is it installed?")

def open_mac_app(app_name):
    apps = {
        "textedit": "open -a TextEdit",
        "calculator": "open -a Calculator"
    }
    if app_name in apps:
        try:
            process = subprocess.Popen(apps[app_name], shell=True)
            opened_apps[app_name] = process
            speak(f"Opened {app_name}")
        except subprocess.CalledProcessError:
            speak(f"Couldn't open {app_name}")
    else:
        speak(f"Don't know how to open {app_name} on macOS.")

def open_linux_app(app_name):
    apps = {
        "gedit": "gedit",
        "calculator": "gnome-calculator"
    }
    if app_name in apps:
        try:
            process = subprocess.Popen(apps[app_name], shell=True)
            opened_apps[app_name] = process
            speak(f"Opened {app_name}")
        except subprocess.CalledProcessError:
            speak(f"Couldn't open {app_name}")
    else:
        speak(f"Don't know how to open {app_name} on Linux.")

def close_app(app_name):
    if app_name in opened_apps:
        process = opened_apps[app_name]
        if process is not None:
            process.terminate()
        speak(f"Closed {app_name}")
        del opened_apps[app_name]
    else:
        speak(f"{app_name} wasn't opened by me.")

def list_open_apps():
    if opened_apps:
        speak("Opened apps:")
        for app in opened_apps:
            speak(app)
    else:
        speak("No apps opened by me are running.")

def track_all_open_apps():
    open_apps = []
    for process in psutil.process_iter(['name']):
        try:
            name = process.info['name']
            if name and name.lower() not in ['system', 'registry', 'svchost.exe']:
                open_apps.append(name)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    if open_apps:
        speak("Active apps:")
        for app in open_apps[:10]:  # limit to 10 for speed
            speak(app)
    else:
        speak("No active apps found.")

# ---------------------- DICTIONARY -----------------------

def get_meaning(word):
    synsets = wordnet.synsets(word)
    if synsets:
        definition = synsets[0].definition()
        speak(f"{word} means: {definition}")
    else:
        speak(f"No meaning found for {word}.")

# ---------------------- MAIN LOOP -----------------------

if __name__ == "__main__":
    speak("Hello! I’m ready. Say open, close, define, or stop.")

    while True:
        speak("What would you like to do?")
        command = listen()

        if not command:
            continue

        if "open" in command:
            app_name = command.replace("open", "").strip()
            open_app(app_name)

        elif "close" in command:
            app_name = command.replace("close", "").strip()
            close_app(app_name)

        elif "list open apps" in command:
            list_open_apps()

        elif "track open apps" in command:
            track_all_open_apps()

        elif "define" in command or "meaning of" in command:
            word = command.replace("define", "").replace("meaning of", "").strip()
            get_meaning(word)

        elif command in ["exit", "quit", "stop"]:
            speak("Goodbye!")
            break

        else:
            speak("I didn’t understand. Try open, close, define, or stop.")
