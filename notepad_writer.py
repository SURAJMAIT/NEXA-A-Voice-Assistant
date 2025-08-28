import os
import pyttsx3
import speech_recognition as sr
import pyautogui
import time
import subprocess
import win32com.client as win32

# Initialize engines
recognizer = sr.Recognizer()
engine = pyttsx3.init()
interactive_notepad_active = False
interactive_word_active = False

def speak(text):
    engine.say(text)
    engine.runAndWait()

def listen():
    try:
        with sr.Microphone() as source:
            print("Listening...")
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=8)
        command = recognizer.recognize_google(audio)
        print(f"You said: {command}")
        return command.lower()
    except sr.UnknownValueError:
        speak("Sorry, I didn't catch that.")
    except sr.RequestError:
        speak("Speech service is unavailable.")
    except Exception as e:
        print(f"Listening error: {e}")
    return ""

# NOTEPAD FUNCTIONS

def open_notepad():
    subprocess.Popen(["notepad.exe"])
    speak("Opening Notepad.")
    time.sleep(2)

def write_to_notepad(text):
    pyautogui.typewrite(text, interval=0.05)
    speak("Written successfully.")

def save_notepad():
    pyautogui.hotkey('ctrl', 's')
    time.sleep(1)
    speak("What name should I save the file as?")
    filename = listen().replace(" ", "_") + ".txt"
    pyautogui.typewrite(filename)
    pyautogui.hotkey('enter')
    speak(f"File saved as {filename}")
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    speak("Notepad closed.")

def open_notepad_file():
    subprocess.Popen(["notepad.exe"])
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'o')
    time.sleep(1)
    speak("Tell me the file name you want to open.")
    filename = listen().replace(" ", "_") + ".txt"
    pyautogui.typewrite(filename)
    pyautogui.hotkey('enter')
    speak(f"Opened {filename}")

def close_notepad():
    pyautogui.hotkey('alt', 'f4')
    speak("Notepad closed.")

# WORD FUNCTIONS

def open_word(filename=None):
    word = win32.Dispatch('Word.Application')
    word.Visible = True
    if filename:
        full_path = os.path.abspath(filename)
        word.Documents.Open(full_path)
        speak(f"Opened {filename}")
    else:
        word.Documents.Add()
        speak("New Word document created.")
    return word

def write_to_word(text, word_app):
    word_app.Selection.TypeText(text)
    speak("Written successfully.")

def save_word(word_app):
    speak("What name should I save the Word document as?")
    filename = listen().replace(" ", "_") + ".docx"
    full_path = os.path.abspath(filename)
    word_app.ActiveDocument.SaveAs(full_path)
    speak(f"Word document saved as {filename}")
    word_app.Quit()
    speak("Word closed.")

def open_word_file():
    speak("Tell me the Word document name you want to open.")
    filename = listen().replace(" ", "_") + ".docx"
    return open_word(filename)

def close_word(word_app):
    word_app.Quit()
    speak("Word closed.")

# NOTES FUNCTION

def open_notes():
    try:
        subprocess.Popen([
            "C:\\Program Files\\WindowsApps\\Microsoft.WindowsStickyNotes_8.0.0.0_x64__8wekyb3d8bbwe\\StickyNotes.exe"
        ])
        speak("Opening Sticky Notes.")
    except Exception as e:
        speak(f"Failed to open Sticky Notes: {e}")

# INTERACTIVE WRITERS

def interactive_notepad_writer():
    global interactive_notepad_active
    open_notepad()
    interactive_notepad_active = True

    while True:
        speak("What should I write?")
        text = listen()
        if text:
            pyautogui.typewrite(text + "\n", interval=0.05)

        speak("Do you want to write more?")
        response = listen()

        if "yes" in response:
            continue
        elif "no" in response:
            speak("Okay. Writing session finished.")
            break
        else:
            speak("I didn't understand that. Please say yes or no.")
            response = listen()
            if "yes" in response:
                continue
            else:
                speak("No clear answer. Ending session.")
                break

def interactive_word_writer(word_app):
    global interactive_word_active
    interactive_word_active = True

    while True:
        speak("What should I write?")
        text = listen()
        if text:
            write_to_word(text + "\n", word_app)

        speak("Do you want to write more?")
        response = listen()

        if "yes" in response:
            continue
        elif "no" in response:
            speak("Okay. Writing session finished.")
            break
        else:
            speak("I didn't understand that. Please say yes or no.")
            response = listen()
            if "yes" in response:
                continue
            else:
                speak("No clear answer. Ending session.")
                break

# COMMAND EXECUTION

def execute_command(command, word_app=None):
    global current_mode, interactive_notepad_active, interactive_word_active

    if "write text in notepad" in command:
        interactive_notepad_writer()
        current_mode = "interactive_notepad"

    elif "write text in word" in command:
        word_app = open_word()
        interactive_word_writer(word_app)
        current_mode = "interactive_word"

    elif "open notepad" in command:
        open_notepad()
        current_mode = "notepad"

    elif "open text file" in command:
        open_notepad_file()
        current_mode = "notepad"

    elif "open word file" in command:
        word_app = open_word_file()
        current_mode = "word"

    elif "open word" in command:
        word_app = open_word()
        current_mode = "word"

    elif "open notes" in command:
        open_notes()
        current_mode = "notes"

    elif "write text" == command.strip():
        speak("What should I write?")
        text = listen()
        if current_mode == "notepad":
            write_to_notepad(text)
        elif current_mode == "word" and word_app:
            write_to_word(text, word_app)
        else:
            speak("No writing application open.")

    elif "save document" in command:
        if current_mode == "notepad" or current_mode == "interactive_notepad":
            save_notepad()
            current_mode = None
            interactive_notepad_active = False
        elif current_mode == "word" or current_mode == "interactive_word":
            if word_app:
                save_word(word_app)
                current_mode = None
                interactive_word_active = False
        else:
            speak("Nothing to save.")

    elif "close notepad" in command and current_mode == "notepad":
        close_notepad()
        current_mode = None

    elif "close word" in command and current_mode == "word":
        close_word(word_app)
        current_mode = None

    else:
        speak("Command not recognized.")

    return word_app

# MAIN LOOP

def main():
    global current_mode
    current_mode = None
    word_app = None

    speak("Hello! I'm ready to assist you with Notepad, Word, and Notes.")

    while True:
        command = listen()
        if command:
            if "exit" in command:
                speak("Goodbye!")
                break
            word_app = execute_command(command, word_app)
        time.sleep(0.5)

if __name__ == "__main__":
    main()
