import speech_recognition as sr
import pyttsx3
import platform
import subprocess
import psutil
import os
import webbrowser

# Initialize recognizer and text-to-speech engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()

# Dictionary to keep track of opened applications
opened_apps = {}

def speak(text):
    engine.say(text)
    engine.runAndWait()

def listen():
    with sr.Microphone() as source:
        print("Listening...")
        audio = recognizer.listen(source)
    try:
        command = recognizer.recognize_google(audio)
        print(f"You said: {command}")
    except sr.UnknownValueError:
        speak("Sorry, I did not get that.")
        return ""
    return command.lower()

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
                os.startfile(app_exe)  # for URI-based apps
            elif app_exe.endswith(".exe"):
                subprocess.Popen(app_exe)
            else:
                webbrowser.open(app_exe)  # fallback
            opened_apps[app_name] = None
            speak(f"Opening {app_name}")
        except Exception as e:
            speak(f"Sorry, I couldn't open {app_name}. Error: {e}")
    else:
        try:
            subprocess.Popen([app_name])
            opened_apps[app_name] = None
            speak(f"Opening {app_name}")
        except FileNotFoundError:
            speak(f"Sorry, I couldn't find {app_name}. Please make sure it's installed.")

def open_mac_app(app_name):
    apps = {
        "textedit": "open -a TextEdit",
        "calculator": "open -a Calculator"
    }
    if app_name in apps:
        try:
            process = subprocess.Popen(apps[app_name], shell=True)
            opened_apps[app_name] = process
            speak(f"Opening {app_name}")
        except subprocess.CalledProcessError:
            speak(f"Sorry, I couldn't open {app_name}")
    else:
        speak(f"I don't know how to open {app_name} on macOS.")

def open_linux_app(app_name):
    apps = {
        "gedit": "gedit",
        "calculator": "gnome-calculator"
    }
    if app_name in apps:
        try:
            process = subprocess.Popen(apps[app_name], shell=True)
            opened_apps[app_name] = process
            speak(f"Opening {app_name}")
        except subprocess.CalledProcessError:
            speak(f"Sorry, I couldn't open {app_name}")
    else:
        speak(f"I don't know how to open {app_name} on Linux.")

def close_app(app_name):
    if app_name in opened_apps:
        process = opened_apps[app_name]
        if process is not None:
            process.terminate()
        speak(f"Closing {app_name}")
        del opened_apps[app_name]
    else:
        speak(f"I don't have a record of opening {app_name}")

def list_open_apps():
    if opened_apps:
        speak("The following applications are open:")
        for app in opened_apps:
            speak(app)
    else:
        speak("No applications are currently open.")

def track_all_open_apps():
    open_apps = []
    for process in psutil.process_iter(['name']):
        try:
            if process.name() not in ['System', 'Registry', 'svchost.exe']:
                open_apps.append(process.name())
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    if open_apps:
        speak("The following applications are currently open:")
        for app in open_apps:
            speak(app)
    else:
        speak("No active user applications are currently running.")

if __name__ == "__main__":
    speak("Hello! I'm ready to help you open and manage applications.")
    while True:
        speak("What would you like to do?")
        command = listen()
        if command:
            if "open" in command:
                app_name = command.replace("open", "").strip()
                speak(f"Trying to open {app_name}")
                open_app(app_name)
            elif "close" in command:
                app_name = command.replace("close", "").strip()
                speak(f"Trying to close {app_name}")
                close_app(app_name)
            elif "list open apps" in command:
                list_open_apps()
            elif "track open apps" in command:
                track_all_open_apps()
            elif "exit" in command or "quit" in command:
                speak("Goodbye!")
                break
            else:
                speak("I didn't understand that command. You can say 'open', 'close', 'list open apps', 'track open apps', or 'exit'.")
        else:
            speak("I didn't catch that. Could you please repeat?")
