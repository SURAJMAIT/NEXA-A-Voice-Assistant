import sys
import speech_recognition as sr
import pyttsx3
import platform
import subprocess
import psutil
import os
import webbrowser
import time
from datetime import datetime, timedelta # Corrected import for datetime and added timedelta for stopwatch
from dateutil import parser
import re
import requests
from nltk.corpus import wordnet
import nltk
import keyboard
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, WebDriverException, TimeoutException
import pygetwindow as gw # Used by YouTube controller (Windows-specific)
import threading # Keep threading for reminders and stopwatch

# --- Global Initializations ---
# Initialize recognizer and text-to-speech engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()
engine.setProperty('rate', 160) # Set a consistent speech rate

# Lock for thread-safe speech (used in reminder and stopwatch)
speak_lock = threading.Lock()

# Dictionary to keep track of opened applications (from app_launcher)
# Note: Tracking opened apps this way is not foolproof, especially for apps
# not opened directly by subprocess.Popen.
opened_apps = {}

# Global variables for Notepad/Word interaction (from notepad_writer)
interactive_notepad_active = False
interactive_word_active = False
current_mode = None # To track the current active mode (notepad, word, etc.)
word_app = None # To hold the Word application instance

# Global variable for Web Navigator (from web_navigator)
driver = None # To hold the Selenium webdriver instance

# Global variables for Stopwatch
stopwatch_running = False
stopwatch_start_time = None
stopwatch_stop_event = threading.Event() 

stopwatch_thread = None

# Download wordnet dataset (from dictionary)
try:
    nltk.data.find('corpora/wordnet')
except LookupError:
    print("NLTK wordnet data not found. Attempting to download...")
    try:
        nltk.download('wordnet')
        print("NLTK wordnet data downloaded successfully.")
    except Exception as e:
        print(f"Failed to download NLTK wordnet data: {e}")
        print("Please try running 'import nltk; nltk.download(\'wordnet\')' in a separate Python session.")


# --- Core Voice Functions (Consolidated) ---

def speak(text):
    """Speaks the given text, using a lock for thread safety."""
    with speak_lock:
        print(f"Assistant: {text}")
        engine.say(text)
        engine.runAndWait()

def listen(timeout=5, phrase_time_limit=8, prompt=None): # Added prompt parameter
    """Listens for a command and returns recognized speech."""
    try:
        with sr.Microphone() as source:
            print("Listening...")
            recognizer.adjust_for_ambient_noise(source, duration=1) # Adjust for ambient noise
            if prompt: # Speak the prompt if provided
                speak(prompt)
                print(f"Assistant (listening): {prompt}")
            audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
        command = recognizer.recognize_google(audio)
        print(f"You said: {command}")
        return command.lower().strip()
    except sr.UnknownValueError:
        print("Sorry, I didn't catch that.")
        return ""
    except sr.RequestError:
        speak("Speech service is unavailable. Check your connection.")
        return ""
    except Exception as e:
        print(f"Listening error: {e}")
        return ""

# --- Module-Specific Functions (Copied and Integrated) ---

# --- App Launcher Functions ---
def open_app(app_name):
    """Opens a specified application based on the operating system."""
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
    """Opens a Windows application."""
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
                # Use os.startfile for URI schemes on Windows
                os.startfile(app_exe)
                speak(f"Opening {app_name}")
                # Note: os.startfile doesn't return a process object to track easily
                opened_apps[app_name] = None
            elif app_exe.endswith(".exe"):
                # Use subprocess.Popen for executables to get a process object
                process = subprocess.Popen(app_exe)
                opened_apps[app_name] = process # Store the process object
                speak(f"Opening {app_name}")
            else:
                # Fallback to webbrowser for other cases
                webbrowser.open(app_exe)
                speak(f"Opening {app_name}")
                opened_apps[app_name] = None
        except FileNotFoundError:
            speak(f"Sorry, I couldn't find the executable for {app_name}.")
            if app_name in opened_apps:
                del opened_apps[app_name]
        except Exception as e:
            speak(f"Sorry, I couldn't open {app_name}. Error: {e}")
            if app_name in opened_apps:
                del opened_apps[app_name]
    else:
        # Try opening as a general command
        try:
            process = subprocess.Popen([app_name])
            opened_apps[app_name] = process # Store the process object
            speak(f"Opening {app_name}")
        except FileNotFoundError:
            speak(f"Sorry, I couldn't find {app_name}. Please make sure it's installed.")
            if app_name in opened_apps:
                del opened_apps[app_name]
        except Exception as e:
             speak(f"Sorry, I couldn't open {app_name}. Error: {e}")
             if app_name in opened_apps:
                del opened_apps[app_name]


def open_mac_app(app_name):
    """Opens a macOS application."""
    apps = {
        "textedit": "TextEdit", # Use application name for 'open -a'
        "calculator": "Calculator"
    }
    if app_name in apps:
        try:
            # Use 'open -a' to open applications on macOS
            process = subprocess.Popen(["open", "-a", apps[app_name]])
            opened_apps[app_name] = process # Store the process object
            speak(f"Opening {app_name}")
        except FileNotFoundError:
            speak(f"Sorry, I couldn't find the application {app_name}.")
            if app_name in opened_apps:
                del opened_apps[app_name]
        except Exception as e:
             speak(f"Sorry, I couldn't open {app_name}. Error: {e}")
             if app_name in opened_apps:
                del opened_apps[app_name]
    else:
        speak(f"I don't know how to open {app_name} on macOS.")

def open_linux_app(app_name):
    """Opens a Linux application."""
    apps = {
        "gedit": "gedit",
        "calculator": "gnome-calculator"
    }
    if app_name in apps:
        try:
            # Use subprocess.Popen for Linux commands
            process = subprocess.Popen([apps[app_name]])
            opened_apps[app_name] = process # Store the process object
            speak(f"Opening {app_name}")
        except FileNotFoundError:
            speak(f"Sorry, I couldn't find the command for {app_name}.")
            if app_name in opened_apps:
                del opened_apps[app_name]
        except Exception as e:
             speak(f"Sorry, I couldn't open {app_name}. Error: {e}")
             if app_name in opened_apps:
                del opened_apps[app_name]
    else:
        speak(f"I don't know how to open {app_name} on Linux.")


def close_app(app_name):
    """Closes a previously opened application."""
    app_name_lower = app_name.lower()
    closed = False

    # Attempt to close using the stored process object first
    if app_name in opened_apps and opened_apps[app_name] is not None:
        process = opened_apps[app_name]
        try:
            print(f"Attempting to terminate process object for {app_name}...")
            process.terminate()
            speak(f"Closing {app_name}")
            closed = True
        except psutil.NoSuchProcess:
            print(f"Process object for {app_name} not found (might already be closed).")
            speak(f"{app_name} process not found (might already be closed).")
            closed = True # Consider it closed if process is gone
        except Exception as e:
            print(f"Error terminating process object for {app_name}: {e}")

    # If not closed via process object or if process object was None, try other methods
    if not closed:
        system = platform.system()
        if system == "Windows":
            # Use taskkill on Windows as a more reliable method
            try:
                # '/IM' specifies the image name (executable name)
                # '/F' forces termination
                print(f"Attempting to close {app_name_lower}.exe with taskkill...")
                # Use shell=True for taskkill to work reliably with paths/quotes if needed, but be cautious
                result = subprocess.run(["taskkill", "/IM", f"{app_name_lower}.exe", "/F"], check=True, capture_output=True, text=True)
                print(f"taskkill stdout: {result.stdout}")
                print(f"taskkill stderr: {result.stderr}")
                speak(f"Closing {app_name} using taskkill.")
                closed = True
            except subprocess.CalledProcessError as e:
                # taskkill returns non-zero if process not found or other error
                print(f"taskkill failed for {app_name}: {e.stderr}")
                speak(f"Could not find a running process named {app_name} to close using taskkill.")
            except FileNotFoundError:
                speak("taskkill command not found. Cannot force close applications.")
            except Exception as e:
                speak(f"An error occurred while trying to close {app_name}: {e}")

        # Fallback for other OS or if taskkill failed
        if not closed:
            # Iterate through all processes and try to match by name or executable path
            print(f"Attempting to close {app_name_lower} using psutil iteration...")
            for proc in psutil.process_iter(['name', 'exe']):
                try:
                    proc_name = proc.name().lower()
                    proc_exe = proc.exe().lower() if proc.exe() else ""
                    # Check if the app name is a substring of the process name or executable path
                    if app_name_lower in proc_name or app_name_lower in proc_exe:
                        print(f"Found matching process: {proc.pid} - {proc.name()}")
                        proc.terminate()
                        speak(f"Closing {app_name}")
                        closed = True
                        break # Assume the first match is the correct one
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass # Ignore processes we can't access or that are already gone
                except Exception as e:
                    print(f"Error terminating process {proc.name()} during psutil iteration: {e}")
                    speak(f"Could not close {app_name} due to an error during process iteration.")
                    closed = True # Consider attempted

    if closed and app_name in opened_apps:
         del opened_apps[app_name] # Remove from tracking if we attempted to close it
    elif not closed:
         speak(f"Could not find a running process for {app_name} to close.")


def list_open_apps():
    """Lists applications tracked by the app launcher."""
    if opened_apps:
        speak("The following applications are open:")
        for app in opened_apps:
            speak(app)
    else:
        speak("No applications are currently open that I launched.")

def track_all_open_apps():
    """Lists all currently running user applications."""
    open_apps = []
    for process in psutil.process_iter(['name']):
        try:
            # Exclude common system processes
            if process.name() not in ['System', 'Registry', 'svchost.exe', 'RuntimeBroker.exe', 'explorer.exe', 'conhost.exe']:
                open_apps.append(process.name())
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    if open_apps:
        speak("The following applications are currently running:")
        # Speak a reasonable number to avoid a very long list
        # Sort for consistency
        open_apps.sort()
        for app in open_apps[:20]: # Limit to first 20 for brevity
            speak(app)
        if len(open_apps) > 20:
            speak("and more.")
    else:
        speak("No active user applications are currently running.")


# --- Dictionary Functions ---
def get_meaning(word):
    """Gets and speaks the meaning of a word using WordNet."""
    synsets = wordnet.synsets(word)
    if synsets:
        definition = synsets[0].definition()
        response = f"The meaning of {word} is: {definition}"
    else:
        response = f"Sorry, I couldn't find the meaning of {word}."

    print(response)
    speak(response)


# --- Notepad/Word Writer Functions ---
# NOTE: These functions heavily rely on PyAutoGUI and win32com, which are Windows-specific.
# They may not work on other operating systems.

def open_notepad():
    """Opens Notepad."""
    if platform.system() == "Windows":
        try:
            subprocess.Popen(["notepad.exe"])
            speak("Opening Notepad.")
            time.sleep(2) # Give time for the application to open
        except FileNotFoundError:
             speak("Notepad executable not found.")
        except Exception as e:
             speak(f"Failed to open Notepad: {e}")
    else:
        speak("Notepad functions are only supported on Windows.")

def write_to_notepad(text):
    """Types text into the currently active window (assumed to be Notepad)."""
    if platform.system() == "Windows":
        try:
            pyautogui.typewrite(text, interval=0.05)
            speak("Written successfully.")
        except Exception as e:
            speak(f"Error writing to Notepad: {e}")
    else:
         speak("Writing functions are only supported on Windows.")

def save_notepad():
    """Saves the current Notepad file."""
    if platform.system() == "Windows":
        try:
            pyautogui.hotkey('ctrl', 's')
            time.sleep(1)
            speak("What name should I save the file as?")
            filename = listen().replace(" ", "_") + ".txt"
            if not filename.endswith(".txt"): # Ensure .txt extension
                filename += ".txt"
            pyautogui.typewrite(filename)
            pyautogui.hotkey('enter')
            speak(f"File saved as {filename}")
            time.sleep(1)
            # pyautogui.hotkey('alt', 'f4') # Optionally close after saving
            # speak("Notepad closed.")
        except Exception as e:
            speak(f"Error saving Notepad file: {e}")
    else:
        speak("Saving functions are only supported on Windows.")

def open_notepad_file():
    """Opens a specific text file in Notepad."""
    if platform.system() == "Windows":
        try:
            open_notepad() # Open notepad first
            pyautogui.hotkey('ctrl', 'o')
            time.sleep(1)
            speak("Tell me the file name you want to open.")
            filename = listen().replace(" ", "_")
            if not filename.endswith(".txt"): # Assume .txt if no extension given
                filename += ".txt"
            pyautogui.typewrite(filename)
            pyautogui.hotkey('enter')
            speak(f"Opened {filename}")
        except Exception as e:
            speak(f"Error opening Notepad file: {e}")
    else:
        speak("Opening files in Notepad is only supported on Windows.")


def close_notepad():
    """Closes Notepad."""
    if platform.system() == "Windows":
        try:
            pyautogui.hotkey('alt', 'f4')
            speak("Notepad closed.")
        except Exception as e:
            speak(f"Error closing Notepad: {e}")
    else:
        speak("Closing Notepad is only supported on Windows.")

def open_word(filename=None):
    """Opens Microsoft Word, optionally opening a specific file."""
    if platform.system() == "Windows":
        try:
            import win32com.client as win32
            word = win32.Dispatch('Word.Application')
            word.Visible = True
            if filename:
                full_path = os.path.abspath(filename)
                # Check if file exists before trying to open
                if os.path.exists(full_path):
                     word.Documents.Open(full_path)
                     speak(f"Opened {filename}")
                else:
                     speak(f"File not found: {filename}. Creating a new document instead.")
                     word.Documents.Add()
            else:
                word.Documents.Add()
                speak("New Word document created.")
            return word
        except ImportError:
            speak("win32com library not found. Please install it for Word functions.")
            return None
        except Exception as e:
            speak(f"Failed to open Word: {e}. Make sure Microsoft Word is installed.")
            return None
    else:
        speak("Microsoft Word functions are only supported on Windows.")
        return None

def write_to_word(text, word_app):
    """Types text into the active Word document."""
    if platform.system() == "Windows" and word_app:
        try:
            word_app.Selection.TypeText(text)
            speak("Written successfully.")
        except Exception as e:
            speak(f"Error writing to Word: {e}")
    elif platform.system() != "Windows":
         speak("Writing functions are only supported on Windows.")
    else:
        speak("Word application is not open.")


def save_word(word_app):
    """Saves the active Word document."""
    if platform.system() == "Windows" and word_app:
        try:
            speak("What name should I save the Word document as?")
            filename = listen().replace(" ", "_")
            if not filename.endswith(".docx"): # Ensure .docx extension
                 filename += ".docx"
            full_path = os.path.abspath(filename)
            word_app.ActiveDocument.SaveAs(full_path)
            speak(f"Word document saved as {filename}")
            # word_app.Quit() # Optionally quit after saving
            # speak("Word closed.")
        except Exception as e:
            speak(f"Error saving Word document: {e}")
    elif platform.system() != "Windows":
         speak("Saving functions are only supported on Windows.")
    else:
        speak("Word application is not open.")


def open_word_file():
    """Opens a specific Word document."""
    if platform.system() == "Windows":
        speak("Tell me the Word document name you want to open.")
        filename = listen().replace(" ", "_")
        if not filename.endswith(".docx"): # Assume .docx if no extension given
             filename += ".docx"
        return open_word(filename)
    else:
        speak("Opening Word files is only supported on Windows.")
        return None

def close_word(word_app):
    """Closes Microsoft Word."""
    if platform.system() == "Windows" and word_app:
        try:
            word_app.Quit()
            speak("Word closed.")
        except Exception as e:
            speak(f"Error closing Word: {e}")
    elif platform.system() != "Windows":
         speak("Closing Word is only supported on Windows.")
    else:
        speak("Word application is not open.")


def open_notes():
    """Opens Windows Sticky Notes."""
    if platform.system() == "Windows":
        try:
            # This path might vary slightly depending on Windows version and installation
            # Attempt common paths or use shell:StickyNotes
            try:
                subprocess.Popen(["shell:StickyNotes"])
            except FileNotFoundError:
                 # Fallback to a common direct path if shell command fails
                 subprocess.Popen([
                    "C:\\Program Files\\WindowsApps\\Microsoft.MicrosoftStickyNotes_8.0.0.0_x64__8wekyb3d8bbwe\\StickyNotes.exe"
                 ])
            speak("Opening Sticky Notes.")
        except FileNotFoundError:
             speak("Sticky Notes executable not found.")
        except Exception as e:
            speak(f"Failed to open Sticky Notes: {e}")
    else:
        speak("Sticky Notes are only available on Windows.")

# --- Reminder Functions ---

def normalize_compact_time(text):
    """Normalizes time formats like '2pm' to '2:00 pm'."""
    match = re.match(r'^(\d{1,4})\s*(a\.?m\.?|p\.?m\.?)$', text.strip())
    if match:
        digits, meridian = match.groups()
        digits = digits.zfill(4) # Pad with leading zeros if needed (e.g., 2pm -> 0002pm)
        hours = int(digits[:-2])
        minutes = int(digits[-2:])
        return f"{hours}:{minutes:02d} {meridian}"
    return text

def wait_and_remind(reminder_time, message):
    """Waits until the specified time and then speaks the reminder message."""
    while True:
        # Use the imported datetime class directly
        if datetime.now() >= reminder_time:
            speak(f"Reminder: {message}")
            break
        time.sleep(1) # Check every second

def handle_set_reminder():
    """Guides the user through setting a reminder."""
    # Call listen with the prompt argument
    time_input = listen(prompt="When should I remind you?")
    if not time_input:
        speak("No time provided for the reminder.")
        return

    time_input = normalize_compact_time(time_input)

    try:
        # Use the imported datetime class directly
        now = datetime.now()
        # Attempt to parse the time input
        parsed_time = parser.parse(time_input, fuzzy=True) # fuzzy=True allows for more flexible parsing
        reminder_time = parsed_time.replace(year=now.year, month=now.month, day=now.day)

        # If the reminder time is in the past today, set it for the next day
        if reminder_time < now:
            reminder_time = reminder_time.replace(day=now.day + 1)

    except Exception as e:
        speak(f"Sorry, I couldn't understand the time: {e}")
        return

    # Call listen with the prompt argument
    message = listen(prompt="What do you want me to remind you?")
    if not message:
        speak("No message provided for the reminder.")
        return

    formatted_time = reminder_time.strftime('%I:%M %p').lstrip('0')  # e.g., "2:18 AM"
    speak(f"Setting your reminder for {formatted_time}")
    # Run the reminder in a separate thread so it doesn't block the main program
    threading.Thread(target=wait_and_remind, args=(reminder_time, message), daemon=True).start()


# --- Weather Functions ---
# OpenWeatherMap API key (replace with your actual key if needed)
# NOTE: The provided key 'da06cb560539c44265f3b751c6d9f489' might have limitations or expire.
API_KEY = "da06cb560539c44265f3b751c6d9f489"

def get_weather(city):
    """Fetches weather for the given city using OpenWeatherMap API."""
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={API_KEY}&units=metric"
    try:
        response = requests.get(url)
        print(f"🌐 Weather API response code: {response.status_code}")
        if response.status_code == 200:
            data = response.json()
            weather = data["weather"][0]["description"]
            temperature = data["main"]["temp"]
            return f"The weather in {city.title()} is {weather} with a temperature of {temperature} degrees Celsius."
        else:
            return f"Sorry, I couldn't find the weather for {city}."
    except requests.exceptions.RequestException as e:
        print(f"Network error: {e}")
        return "Sorry, I couldn't reach the weather service."


# --- Web Navigator Functions ---
# NOTE: These functions require Selenium and a compatible browser driver (like chromedriver)
# to be installed and in your system's PATH.

def open_browser(browser_name):
    """Opens a specified web browser using Selenium."""
    global driver
    if driver:
        speak("A browser is already open. Closing the current one.")
        try:
            driver.quit()
        except WebDriverException:
            pass # Ignore if driver is already closed or invalid
        driver = None # Reset driver

    try:
        if "chrome" in browser_name:
            speak("Opening Chrome.")
            # Ensure chromedriver is in PATH or specify its location
            driver = webdriver.Chrome()
        elif "firefox" in browser_name:
            speak("Opening Firefox.")
            # Ensure geckodriver is in PATH or specify its location
            driver = webdriver.Firefox()
        elif "edge" in browser_name:
            speak("Opening Edge.")
            # Ensure msedgedriver is in PATH or specify its location
            driver = webdriver.Edge()
        else:
            speak("Browser not supported. Please say Chrome, Firefox, or Edge.")
            return None
        return driver
    except WebDriverException as e:
        speak(f"Failed to open browser: {e}. Make sure the browser and driver are installed correctly.")
        driver = None
        return None

def search_topic(topic):
    """Searches for a topic on Google in the current browser."""
    global driver
    if driver:
        try:
            speak(f"Searching for {topic} on Google.")
            driver.get("https://www.google.com")
            # Use WebDriverWait for robustness
            wait = WebDriverWait(driver, 10)
            search_box = wait.until(EC.presence_of_element_located((By.NAME, "q")))
            search_box.send_keys(topic)
            search_box.send_keys(Keys.RETURN)
        except TimeoutException:
            speak("Search box did not appear in time.")
        except NoSuchElementException:
            speak("Could not find the search box.")
        except WebDriverException as e:
            speak(f"An error occurred during search: {e}")
    else:
        speak("No browser is currently open.")

def scroll_page(direction):
    """Scrolls the current page up or down."""
    global driver
    if driver:
        try:
            if direction == 'up':
                speak("Scrolling up.")
                driver.execute_script("window.scrollBy(0, -500);") # Scroll up by 500 pixels
            elif direction == 'down':
                speak("Scrolling down.")
                driver.execute_script("window.scrollBy(0, 500);") # Scroll down by 500 pixels
            else:
                speak("Invalid scroll direction.")
        except WebDriverException as e:
            speak(f"An error occurred while scrolling: {e}")
    else:
        speak("No browser is currently open.")

def click_link(link_text):
    """Clicks on a link containing the specified text."""
    global driver
    if driver:
        try:
            speak(f"Trying to click on a link with text containing '{link_text}'.")
            # Find elements containing the link text (case-insensitive, partial match)
            # Use WebDriverWait to wait for the link to be clickable
            wait = WebDriverWait(driver, 10)
            # Using a more general XPath that looks for 'a' tags containing the text
            # This might be more robust than relying solely on 'id'
            xpath = f"//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{link_text.lower()}')]"
            print(f"Attempting to find link with XPath: {xpath}")
            link = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            print(f"Found link element: {link.text}")
            link.click()
            speak("Clicked the link.")
        except TimeoutException:
            print(f"Timeout: Link with text containing '{link_text}' did not become clickable in time.")
            speak(f"Link with text containing '{link_text}' did not become clickable in time.")
        except NoSuchElementException:
            print(f"NoSuchElementException: Link with text containing '{link_text}' not found.")
            speak(f"Link with text containing '{link_text}' not found.")
        except WebDriverException as e:
            print(f"WebDriverException while clicking link: {e}")
            speak(f"An error occurred while clicking the link: {e}")
        except Exception as e:
            print(f"An unexpected error occurred while clicking link: {e}")
            speak(f"An unexpected error occurred while clicking the link.")
    else:
        speak("No browser is currently open.")

def go_back():
    """Navigates back in the browser history."""
    global driver
    if driver:
        try:
            speak("Going back.")
            driver.back()
        except WebDriverException as e:
            speak(f"An error occurred while going back: {e}")
    else:
        speak("No browser is currently open.")

def read_aloud_links():
    """Reads aloud the text of links on the current page."""
    global driver
    if driver:
        try:
            speak("Reading aloud links on the page.")
            links = driver.find_elements(By.TAG_NAME, "a")
            link_texts = [link.text for link in links if link.text and link.is_displayed()] # Only visible links with text

            if not link_texts:
                speak("No visible links with text found on the page.")
                return

            speak(f"Found {len(link_texts)} links.")
            index = 0
            # Read aloud all links instead of in batches of 10 for simplicity
            for link_text in link_texts:
                 speak(f"Link {index + 1}: {link_text}")
                 index += 1
                 time.sleep(0.5) # Small pause between reading links


        except WebDriverException as e:
            speak(f"An error occurred while reading the page: {e}")
    else:
        speak("No browser is currently open.")

def close_browser():
    """Closes the current browser instance."""
    global driver
    if driver:
        try:
            driver.quit()
            speak("Browser closed.")
        except WebDriverException:
            pass # Ignore if driver is already closed or invalid
        driver = None
    else:
        speak("No browser is currently open.")


# --- YouTube Controller Functions ---
# NOTE: These functions heavily rely on PyAutoGUI and PyGetWindow, which are Windows-specific.
# They may not work reliably on other operating systems.

def open_youtube():
    """Opens YouTube in the default web browser."""
    speak("Opening YouTube.")
    # Using the actual YouTube URL
    # Using driver.get() instead of webbrowser.open() to open in the current Selenium instance
    global driver
    if driver:
        try:
            driver.get("https://www.youtube.com/") # Use actual YouTube URL
            speak("Opened YouTube in the current browser.")
        except WebDriverException as e:
            speak(f"Could not open YouTube in the current browser: {e}")
            speak("Attempting to open in default browser instead.")
            webbrowser.open("https://www.youtube.com/") # Use actual YouTube URL
            speak("Opened YouTube in your default browser.")
    else:
        webbrowser.open("https://www.youtube.com/") # Use actual YouTube URL
        speak("Opened YouTube in your default browser.")

def is_youtube_video_playing():
    """Checks if a YouTube video is likely playing based on window title."""
    # This is a heuristic and might not be perfectly accurate
    if platform.system() == "Windows":
        for window in gw.getWindowsWithTitle("YouTube"):
            if "-" in window.title: # YouTube video titles usually contain a hyphen
                return True
        return False
    else:
        print("Warning: is_youtube_video_playing is Windows-specific.")
        return False # Cannot reliably check on non-Windows


def search_youtube(query):
    """Searches YouTube for a query and attempts to play the first result."""
    global driver
    if driver:
        try:
            speak(f"Searching YouTube for {query} in the current browser.")
            # Navigate to YouTube search results page using the driver
            # Construct the search URL using the actual YouTube search path
            search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
            driver.get(search_url)

            # Attempt to click the first video link using Selenium
            speak("Attempting to open the first video.")
            wait = WebDriverWait(driver, 10)
            # This XPath is a common pattern for video links on YouTube search results
            # It might need adjustment if YouTube's layout changes
            # Looking for the first link with id 'video-title'
            first_video_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@id='video-title']")))
            first_video_link.click()
            speak("Playing the first video.")

        except TimeoutException:
            speak("YouTube search results or first video link did not load in time.")
        except NoSuchElementException:
            speak("Could not find the search results or first video link on YouTube.")
        except WebDriverException as e:
            speak(f"An error occurred during YouTube search or playback: {e}")
    else:
        speak("No browser is currently open. Opening YouTube search in default browser.")
        # Use the actual YouTube search URL for the default browser
        search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
        webbrowser.open(search_url)


def open_youtube_history():
    """Opens the YouTube history page in the current browser if possible."""
    global driver
    if driver:
        try:
            speak("Opening YouTube history in the current browser.")
            driver.get("https://www.youtube.com/feed/history") # Use actual YouTube history URL
        except WebDriverException as e:
            speak(f"Could not open YouTube history in the current browser: {e}")
            speak("Attempting to open in default browser instead.")
            webbrowser.open("https://www.youtube.com/feed/history") # Use actual YouTube history URL
            speak("Opened YouTube history in your default browser.")
    else:
        webbrowser.open("https://www.youtube.com/feed/history") # Use actual YouTube history URL
        speak("Opening YouTube history in your default browser.")


def open_youtube_subscriptions():
    """Opens the YouTube subscriptions page in the current browser if possible."""
    global driver
    if driver:
        try:
            speak("Opening YouTube subscriptions in the current browser.")
            driver.get("https://www.youtube.com/feed/subscriptions") # Use actual YouTube subscriptions URL
        except WebDriverException as e:
            speak(f"Could not open YouTube subscriptions in the current browser: {e}")
            speak("Attempting to open in default browser instead.")
            webbrowser.open("https://www.youtube.com/feed/subscriptions") # Use actual YouTube subscriptions URL
            speak("Opened YouTube subscriptions in your default browser.")
    else:
        webbrowser.open("https://www.youtube.com/feed/subscriptions") # Use actual YouTube subscriptions URL
        speak("Opening YouTube subscriptions in your default browser.")


def go_youtube_home():
    """Goes to the YouTube home page in the current browser if possible."""
    global driver
    if driver:
        try:
            speak("Going to YouTube Home in the current browser.")
            driver.get("https://www.youtube.com/") # Use actual YouTube home URL
        except WebDriverException as e:
            speak(f"Could not go to YouTube Home in the current browser: {e}")
            speak("Attempting to open in default browser instead.")
            webbrowser.open("https://www.youtube.com/") # Use actual YouTube home URL
            speak("Opened YouTube Home in your default browser.")
    else:
        webbrowser.open("https://www.youtube.com/") # Use actual YouTube home URL
        speak("Opening YouTube Home in your default browser.")


def scroll_youtube_down():
    """Scrolls down on the YouTube page."""
    speak("Scrolling down.")
    if platform.system() == "Windows":
        pyautogui.scroll(-700) # Scroll down by 700 units
    else:
        speak("Scrolling down is Windows-specific.")

def scroll_youtube_up():
    """Scrolls up on the YouTube page."""
    speak("Scrolling up.")
    if platform.system() == "Windows":
        pyautogui.scroll(700) # Scroll up by 700 units
    else:
        speak("Scrolling up is Windows-specific.")

def increase_volume(amount=10):
    """Increases the system volume."""
    speak("Increasing volume.")
    if platform.system() == "Windows":
        for _ in range(amount):
            pyautogui.press('volumeup')
            time.sleep(0.05) # Small delay between presses
    else:
        speak("Volume control is Windows-specific.")

def decrease_volume(amount=10):
    """Decreases the system volume."""
    speak("Decreasing volume.")
    if platform.system() == "Windows":
        for _ in range(amount):
            pyautogui.press('volumedown')
            time.sleep(0.05) # Small delay between presses
    else:
        speak("Volume control is Windows-specific.")

def seek_forward(seconds):
    """Seeks forward in the active video (assumed to be YouTube)."""
    speak(f"Forwarding {seconds} seconds.")
    if platform.system() == "Windows":
        # YouTube uses 'L' key to seek forward 10 seconds
        num_presses = seconds // 10
        for _ in range(num_presses):
            pyautogui.press('l')
            time.sleep(0.1)
    else:
        speak("Seeking forward is Windows-specific.")


def seek_backward(seconds):
    """Seeks backward in the active video (assumed to be YouTube)."""
    speak(f"Rewinding {seconds} seconds.")
    if platform.system() == "Windows":
        # YouTube uses 'J' key to seek backward 10 seconds
        num_presses = seconds // 10
        for _ in range(num_presses):
            pyautogui.press('j')
            time.sleep(0.1)
    else:
        speak("Seeking backward is Windows-specific.")


def play_pause_youtube_video():
    """Plays or pauses the active video (assumed to be YouTube)."""
    speak("Toggling play/pause.")
    if platform.system() == "Windows":
        pyautogui.press('space') # Spacebar toggles play/pause on YouTube
    else:
        speak("Play/pause control is Windows-specific.")


def next_youtube_video():
    """Skips to the next video in a playlist or suggestion."""
    speak("Skipping to next video.")
    if platform.system() == "Windows":
        pyautogui.hotkey('shift', 'n') # Shift + N goes to the next video
    else:
        speak("Skipping to next video is Windows-specific.")


def previous_youtube_video():
    """Goes back to the previous video in a playlist or history."""
    speak("Going back to previous video.")
    if platform.system() == "Windows":
        # Note: This relies on PyAutoGUI and the YouTube window being active.
        # It simulates pressing Shift+P. Reliability can vary.
        pyautogui.hotkey('shift', 'p') # Shift + P goes to the previous video
    else:
        speak("Going back to previous video is Windows-specific and may not be reliable.")


def extract_time_from_command(command):
    """Extracts a number of seconds from a command string."""
    words = command.split()
    for i, word in enumerate(words):
        if word.isdigit():
            # Check if the next word is 'seconds' or similar
            if i + 1 < len(words) and ("second" in words[i+1] or "minute" in words[i+1]):
                 # Simple handling for minutes (assume 1 minute = 60 seconds)
                 if "minute" in words[i+1]:
                     return int(word) * 60
                 return int(word)
            # If no time unit is specified, assume it's seconds
            if i + 1 == len(words) or not words[i+1].isalpha():
                 return int(word)
    return 10 # Default to 10 seconds if no number is found


# --- Stopwatch Functions (Integrated from stopwatch.py) ---

def show_live_timer(start_time, stop_event):
    """Shows the elapsed time of the stopwatch in the console."""
    while not stop_event.is_set():
        elapsed = datetime.now() - start_time
        # Use speak_lock to avoid interfering with other speech output
        with speak_lock:
            # sys.stdout.write(f"\rElapsed Time: {str(elapsed).split('.')[0]}   ")
            # sys.stdout.flush()
            # Instead of writing to stdout, we'll speak the time at intervals or on request
            pass # Live console output is tricky with voice input, will rely on 'tell me the time' command
        time.sleep(1)


def start_stopwatch():
    """Starts the stopwatch."""
    global stopwatch_running, stopwatch_start_time, stopwatch_stop_event, stopwatch_thread
    if not stopwatch_running:
        stopwatch_start_time = datetime.now()
        stopwatch_running = True
        stopwatch_stop_event.clear() # Clear the event for a new run
        # Start the live timer thread
        stopwatch_thread = threading.Thread(target=show_live_timer, args=(stopwatch_start_time, stopwatch_stop_event), daemon=True)
        stopwatch_thread.start()
        speak(f"Stopwatch started at: {stopwatch_start_time.strftime('%H:%M:%S')}")
    else:
        speak("Stopwatch is already running.")

def stop_stopwatch():
    """Stops the stopwatch and reports the total duration."""
    global stopwatch_running, stopwatch_start_time, stopwatch_stop_event, stopwatch_thread
    if stopwatch_running:
        stopwatch_stop_event.set() # Signal the thread to stop
        if stopwatch_thread and stopwatch_thread.is_alive():
            stopwatch_thread.join() # Wait for the thread to finish
        end_datetime = datetime.now()
        total_time = end_datetime - stopwatch_start_time
        stopwatch_running = False
        stopwatch_start_time = None # Reset start time
        speak(f"Stopwatch stopped at: {end_datetime.strftime('%H:%M:%S')}")
        speak(f"Total Duration: {str(total_time).split('.')[0]}")
    else:
        speak("Stopwatch is not running.")

def get_stopwatch_time():
    """Tells the current elapsed time if the stopwatch is running."""
    global stopwatch_running, stopwatch_start_time
    if stopwatch_running and stopwatch_start_time:
        elapsed = datetime.now() - stopwatch_start_time
        speak(f"Elapsed time: {str(elapsed).split('.')[0]}")
    elif stopwatch_running and not stopwatch_start_time:
        speak("Stopwatch is running, but the start time was not recorded.")
    else:
        speak("Stopwatch is not running.")


# --- Central Command Processing ---

def process_command(command):
    """Processes the recognized voice command and routes it to the appropriate function."""
    global current_mode, word_app, driver, stopwatch_running

    # --- General Commands ---
    if "exit" in command or "quit" in command or "goodbye" in command:
        speak("Goodbye!")
        # Clean up resources before exiting
        if driver:
            try:
                driver.quit()
            except WebDriverException:
                pass
        # Stop stopwatch thread if running
        if stopwatch_running:
            stopwatch_stop_event.set()
            if stopwatch_thread and stopwatch_thread.is_alive():
                 stopwatch_thread.join()
        sys.exit() # Exit the program

    # --- App Launcher Commands ---
    # Added more specific checks to avoid conflicts with other commands
    elif command.startswith("open ") and "browser" not in command and "notepad" not in command and "word" not in command and "notes" not in command and "youtube" not in command and "file" not in command:
        app_name = command.replace("open", "").strip()
        if app_name:
            speak(f"Trying to open {app_name}")
            open_app(app_name)
        else:
            speak("Please specify which application to open.")
    elif command.startswith("close ") and "browser" not in command and "notepad" not in command and "word" not in command:
         app_name = command.replace("close", "").strip()
         if app_name:
             speak(f"Trying to close {app_name}")
             close_app(app_name)
         else:
             speak("Please specify which application to close.")
    elif "list open apps" in command:
        list_open_apps()
    elif "track open apps" in command:
        track_all_open_apps()

    # --- Dictionary Commands ---
    elif "define" in command or "meaning of" in command:
        word_to_define = command.replace("define", "").replace("meaning of", "").strip()
        if word_to_define:
            get_meaning(word_to_define)
        else:
            speak("Please say the word you want me to define.")

    # --- Notepad/Word Writer Commands ---
    elif "open notepad" in command:
        open_notepad()
        current_mode = "notepad"
    elif "open text file" in command:
        open_notepad_file()
        current_mode = "notepad"
    elif "open word file" in command:
        word_app = open_word_file()
        if word_app:
            current_mode = "word"
    elif "open word" in command:
        word_app = open_word()
        if word_app:
            current_mode = "word"
    elif "open notes" in command:
        open_notes()
        current_mode = "notes" # Track notes as a mode, although no specific notes functions are integrated yet

    elif "write text in notepad" in command:
        # Start interactive notepad session
        open_notepad()
        current_mode = "interactive_notepad"
        speak("Starting interactive notepad session. What should I write? Say 'stop writing' to end.")
        while current_mode == "interactive_notepad":
            text_to_write = listen()
            if text_to_write:
                if "stop writing" in text_to_write:
                    speak("Ending interactive notepad session.")
                    current_mode = "notepad" # Go back to notepad mode
                    break
                write_to_notepad(text_to_write + "\n") # Add newline for readability
            else:
                speak("Didn't catch that. Say 'stop writing' to end.")

    elif "write text in word" in command:
        # Start interactive word session
        if word_app is None:
             word_app = open_word()
             if word_app is None: # If opening word failed
                 speak("Could not open Word to start writing session.")
                 return

        current_mode = "interactive_word"
        speak("Starting interactive word session. What should I write? Say 'stop writing' to end.")
        while current_mode == "interactive_word":
            text_to_write = listen()
            if text_to_write:
                if "stop writing" in text_to_write:
                    speak("Ending interactive word session.")
                    current_mode = "word" # Go back to word mode
                    break
                write_to_word(text_to_write + "\n", word_app) # Add newline for readability
            else:
                 speak("Didn't catch that. Say 'stop writing' to end.")


    elif "write text" in command and current_mode in ["notepad", "word", "interactive_notepad", "interactive_word"]:
        speak("What should I write?")
        text = listen()
        if text:
            if current_mode in ["notepad", "interactive_notepad"]:
                write_to_notepad(text)
            elif current_mode in ["word", "interactive_word"] and word_app:
                write_to_word(text, word_app)
            else:
                speak("No active writing application recognized.")
        else:
            speak("No text provided to write.")

    elif "save document" in command:
        if current_mode in ["notepad", "interactive_notepad"]:
            save_notepad()
            # current_mode = None # Optionally reset mode after saving and closing
            # interactive_notepad_active = False
        elif current_mode in ["word", "interactive_word"] and word_app:
            save_word(word_app)
            # current_mode = None # Optionally reset mode after saving and closing
            # interactive_word_active = False
        else:
            speak("Nothing to save in the current mode.")

    elif "close notepad" in command and current_mode in ["notepad", "interactive_notepad"]:
        close_notepad()
        current_mode = None
        interactive_notepad_active = False
    elif "close word" in command and current_mode in ["word", "interactive_word"] and word_app:
        close_word(word_app)
        current_mode = None
        interactive_word_active = False

    # --- Reminder Commands ---
    elif "set a reminder" in command:
        handle_set_reminder()

    # --- Weather Commands ---
    elif "weather in" in command:
        city = command.replace("weather in", "").strip()
        if city:
            weather_report = get_weather(city)
            speak(weather_report)
        else:
            speak("Please specify a city for the weather report.")
    elif "get weather" in command:
        speak("Which city's weather would you like to know?")
        city = listen()
        if city:
            weather_report = get_weather(city)
            speak(weather_report)
        else:
            speak("No city name provided.")


    # --- Web Navigator Functions ---
    elif "open browser" in command:
        browser_name = command.replace("open browser", "").strip()
        if browser_name:
            open_browser(browser_name)
        else:
            speak("Please specify which browser to open.")
    elif "search for" in command and driver:
        topic = command.replace("search for", "").strip()
        if topic:
            search_topic(topic)
        else:
            speak("Please specify what you want to search for.")
    elif "scroll down" in command and driver:
        scroll_page("down")
    elif "scroll up" in command and driver:
        scroll_page("up")
    elif "click" in command and driver:
        link_text = command.replace("click", "").strip()
        if link_text:
            click_link(link_text)
        else:
            speak("Please specify the text of the link to click.")
    elif "go back" in command and driver:
        go_back()
    elif "read links" in command and driver:
        read_aloud_links()
    elif "close browser" in command:
        close_browser()

    # --- YouTube Controller Commands ---
    elif "open youtube" in command:
        open_youtube()
    elif "search youtube for" in command:
        query = command.replace("search youtube for", "").strip()
        if query:
            search_youtube(query)
        else:
            speak("Please specify what you want to search for on YouTube.")
    # Note: These commands rely on PyAutoGUI/PyGetWindow and are Windows-specific.
    # Their reliability depends on the YouTube window being active and focused.
    elif "play video" in command or "pause video" in command or "play" in command or "pause" in command:
        play_pause_youtube_video()
    elif "scroll down youtube" in command:
        scroll_youtube_down()
    elif "scroll up youtube" in command:
        scroll_youtube_up()
    elif "volume up" in command:
        increase_volume()
    elif "volume down" in command:
        decrease_volume()
    elif "go forward" in command or "forward video" in command:
        seconds = extract_time_from_command(command)
        seek_forward(seconds)
    elif "go backward" in command or "rewind video" in command:
        seconds = extract_time_from_command(command)
        seek_backward(seconds)
    elif "next video" in command:
        next_youtube_video()
    elif "previous video" in command:
        previous_youtube_video()
    elif "open youtube history" in command:
        open_youtube_history()
    elif "open youtube subscriptions" in command:
        open_youtube_subscriptions()
    elif "go youtube home" in command:
        go_youtube_home()

    # --- Stopwatch Commands (New) ---
    elif "start stopwatch" in command:
        start_stopwatch()
    elif "stop stopwatch" in command:
        stop_stopwatch()
    elif "tell me the time" in command or "what is the time" in command or "what time is it" in command:
        get_stopwatch_time()


    # --- Unrecognized Command ---
    else:
        speak("I'm sorry, I didn't understand that command.")
        print(f"Unrecognized command: {command}")


# --- Main Execution Loop ---

def main():
    """Main function to initialize and run the voice assistant."""
    speak("Hello! I'm your voice assistant. How can I help you today?")

    # Start the main listening loop
    while True:
        command = listen()
        if command:
            process_command(command)
        time.sleep(0.1) # Small delay to prevent high CPU usage


if __name__ == "__main__":
    # Ensure necessary libraries are installed
    try:
        import speech_recognition as sr
        import pyttsx3
        import psutil
        import pyautogui
        # win32com and pygetwindow are Windows-specific, handle import errors gracefully
        if platform.system() == "Windows":
            import win32com.client as win32
            import pygetwindow as gw
        import nltk
        from dateutil import parser
        import requests
        import keyboard
        from selenium import webdriver
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.common.exceptions import NoSuchElementException, WebDriverException, TimeoutException

    except ImportError as e:
        print(f"Missing required library: {e}")
        print("Please install the necessary libraries using pip:")
        print("pip install SpeechRecognition pyttsx3 psutil pyautogui nltk python-dateutil requests keyboard selenium")
        if platform.system() == "Windows":
             print("For Windows-specific features, you may also need:")
             print("pip install pypiwin32 pygetwindow")
        sys.exit(1)

    # Download NLTK data if not already present
    try:
        nltk.data.find('corpora/wordnet')
    except LookupError:
        print("NLTK wordnet data not found. Attempting to download...")
        try:
            nltk.download('wordnet')
            print("NLTK wordnet data downloaded successfully.")
        except Exception as e:
            print(f"Failed to download NLTK wordnet data: {e}")
            print("Please try running 'import nltk; nltk.download(\'wordnet\')' in a separate Python session.")


    main()