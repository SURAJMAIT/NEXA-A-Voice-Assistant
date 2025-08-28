import speech_recognition as sr
import pyttsx3
from datetime import datetime
from dateutil import parser
import threading
import time
import re

# Initialize text-to-speech engine
engine = pyttsx3.init()
engine.setProperty('rate', 160)

# Lock for thread-safe speech
speak_lock = threading.Lock()

def speak(text):
    with speak_lock:
        print(f"\nAssistant: {text}")
        engine.say(text)
        engine.runAndWait()

def recognize_speech(prompt=None):
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        if prompt:
            print(f"\nAssistant (listening): {prompt}")
            speak(prompt)
        audio = recognizer.listen(source)

        try:
            command = recognizer.recognize_google(audio).lower()
            print(f"You said: {command}")
            return command
        except sr.UnknownValueError:
            speak("Sorry, I didn't understand that.")
            return None
        except sr.RequestError:
            speak("Network error. Please check your connection.")
            return None

def normalize_compact_time(text):
    match = re.match(r'^(\d{1,4})\s*(a\.?m\.?|p\.?m\.?)$', text.strip())
    if match:
        digits, meridian = match.groups()
        digits = digits.zfill(4)
        hours = int(digits[:-2])
        minutes = int(digits[-2:])
        return f"{hours}:{minutes:02d} {meridian}"
    return text

def wait_and_remind(reminder_time, message):
    while True:
        if datetime.now() >= reminder_time:
            speak(f"Reminder: {message}")
            break
        time.sleep(1)

def handle_set_reminder():
    time_input = recognize_speech("When should I remind you?")
    if not time_input:
        return

    time_input = normalize_compact_time(time_input)

    try:
        now = datetime.now()
        parsed_time = parser.parse(time_input)
        reminder_time = parsed_time.replace(year=now.year, month=now.month, day=now.day)
        if reminder_time < now:
            reminder_time = reminder_time.replace(day=now.day + 1)
    except:
        speak("Sorry, I couldn't understand the time.")
        return

    message = recognize_speech("What do you want me to remind you?")
    if not message:
        return

    formatted_time = reminder_time.strftime('%I:%M %p').lstrip('0')  # e.g., "2:18 AM"
    speak(f"Setting your reminder for {formatted_time}")
    threading.Thread(target=wait_and_remind, args=(reminder_time, message), daemon=True).start()

def process_command(command):
    if "set a reminder" in command:
        handle_set_reminder()
    elif "exit" in command:
        speak("Goodbye!")
        exit()
    else:
        speak("Sorry, I didn't recognize that command.")

def main():
    speak("Hello. What can I help you with today?")
    while True:
        command = recognize_speech()
        if command:
            process_command(command)

if __name__ == "__main__":
    main()