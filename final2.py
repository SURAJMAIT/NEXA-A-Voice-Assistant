import speech_recognition as sr
import pyttsx3
import platform
import subprocess
import keyboard
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, WebDriverException
import threading
import time

# Initialize the recognizer, text-to-speech engine, and global variables
recognizer = sr.Recognizer()
engine = pyttsx3.init()
driver = None

def speak(text):
    engine.say(text)
    engine.runAndWait()

def listen():
    try:
        with sr.Microphone() as source:
            print("Listening...")
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)
        command = recognizer.recognize_google(audio)
        print(f"You said: {command}")
        return command.lower()
    except sr.UnknownValueError:
        speak("Sorry, I didn't get that.")
    except sr.RequestError:
        speak("Speech recognition service error.")
    except Exception as e:
        print(f"Listening error: {e}")
    return ""

# Browser control functions
def open_browser(browser_name):
    global driver
    try:
        if browser_name == "chrome":
            driver = webdriver.Chrome()
        elif browser_name == "firefox":
            driver = webdriver.Firefox()
        elif browser_name == "edge":
            driver = webdriver.Edge()
        else:
            speak("Browser not supported.")
            return None
    except WebDriverException as e:
        speak(f"Failed to open browser: {e}")
    return driver

def search_topic(topic):
    global driver
    if driver:
        try:
            driver.get("https://www.google.com")
            search_box = driver.find_element(By.NAME, "q")
            search_box.send_keys(topic)
            search_box.send_keys(Keys.RETURN)
        except WebDriverException as e:
            speak(f"An error occurred during search: {e}")

def scroll_page(direction):
    global driver
    if driver:
        try:
            driver.execute_script(f"window.scrollBy(0, {'-250' if direction == 'up' else '250'});")
        except WebDriverException as e:
            speak(f"An error occurred while scrolling: {e}")

def click_link(link_text):
    global driver
    if driver:
        try:
            link = driver.find_element(By.PARTIAL_LINK_TEXT, link_text)
            link.click()
        except NoSuchElementException:
            speak("Link not found.")
        except WebDriverException as e:
            speak(f"An error occurred while clicking the link: {e}")

def go_back():
    global driver
    if driver:
        try:
            driver.back()
        except WebDriverException as e:
            speak(f"An error occurred while going back: {e}")

def read_aloud():
    global driver
    if driver:
        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text
            speak(body_text)
        except NoSuchElementException:
            speak("Could not find the text to read.")
        except WebDriverException as e:
            speak(f"An error occurred while reading the page: {e}")

# Application launching functions for different OS
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
    apps = {"notepad": "notepad.exe", "calculator": "calc.exe"}
    if app_name in apps:
        try:
            subprocess.Popen(apps[app_name])
            speak(f"Opening {app_name}")
        except FileNotFoundError:
            speak(f"Sorry, I couldn't find {app_name}")

def open_mac_app(app_name):
    apps = {"textedit": "open -a TextEdit", "calculator": "open -a Calculator"}
    if app_name in apps:
        try:
            subprocess.Popen(apps[app_name], shell=True)
            speak(f"Opening {app_name}")
        except subprocess.CalledProcessError:
            speak(f"Sorry, I couldn't open {app_name}")

def open_linux_app(app_name):
    apps = {"gedit": "gedit", "calculator": "gnome-calculator"}
    if app_name in apps:
        try:
            subprocess.Popen(apps[app_name], shell=True)
            speak(f"Opening {app_name}")
        except subprocess.CalledProcessError:
            speak(f"Sorry, I couldn't open {app_name}")

# Keyboard shortcut handling
def execute_command(command):
    if command == "write something":
        keyboard.press_and_release("win+h")
        speak("Executed Win + H (convert to text) command.")
    elif command == "stop":
        keyboard.press_and_release("win+h")
    elif command == "exit":
        keyboard.press_and_release("win+h")
        speak("Exiting the program.")
        return False  # Signal to exit the loop
    else:
        speak("Unknown command.")
    return True

# Main function to continuously listen for commands
def main():
    global driver
    speak("Hello! I'm ready to help you.")
    
    while True:
        command = listen()
        if command:
            if "open browser" in command:
                browser_name = command.replace("open browser", "").strip()
                driver = open_browser(browser_name)
            elif "search" in command and driver:
                topic = command.replace("search", "").strip()
                search_topic(topic)
            elif "scroll up" in command and driver:
                scroll_page("up")
            elif "scroll down" in command and driver:
                scroll_page("down")
            elif "click" in command and driver:
                link_text = command.replace("click", "").strip()
                click_link(link_text)
            elif "go back" in command and driver:
                go_back()
            elif "read" in command and driver:
                read_aloud()
            elif "open" in command:
                app_name = command.replace("open", "").strip()
                open_app(app_name)
            elif command == "write something" or "stop" or "exit" in command:
                if not execute_command(command):
                    break
            else:
                speak("I'm sorry, I didn't understand that command.")
        time.sleep(0.2)  # Short delay to avoid CPU overuse

if __name__ == "__main__":
    main()
