import speech_recognition as sr
import pyautogui
import webbrowser
import time
import pyttsx3
import pygetwindow as gw

# Initialize speech engine
engine = pyttsx3.init()
engine.setProperty('rate', 160)

def speak(text):
    print(f"Assistant: {text}")
    engine.say(text)
    engine.runAndWait()

def open_youtube():
    speak("Opening YouTube")
    webbrowser.open("https://www.youtube.com")
    time.sleep(5)
    # Focus the first tab
    pyautogui.hotkey('ctrl', '1')

def is_video_playing():
    for window in gw.getWindowsWithTitle("YouTube"):
        if "-" in window.title:
            return True
    return False

def recognize_speech():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening for a command...")
        audio = recognizer.listen(source)
        try:
            command = recognizer.recognize_google(audio).lower()
            print(f"Recognized command: {command}")
            return command
        except sr.UnknownValueError:
            speak("Sorry, I didn't understand that.")
        except sr.RequestError:
            speak("Network error. Please check your connection.")
    return None

def search_youtube(query):
    speak(f"Searching YouTube for {query}.")
    pyautogui.hotkey('ctrl', 'l')  # focus address bar
    time.sleep(0.5)
    pyautogui.typewrite(f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}")
    pyautogui.press('enter')
    time.sleep(4)
    # Play first video
    pyautogui.moveTo(300, 400)  # Coordinates might need adjustment
    pyautogui.click()
    time.sleep(3)

def open_history():
    speak("Opening history tab.")
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.typewrite("https://www.youtube.com/feed/history")
    pyautogui.press('enter')

def open_subscriptions():
    speak("Opening subscriptions tab.")
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.typewrite("https://www.youtube.com/feed/subscriptions")
    pyautogui.press('enter')

def open_home():
    speak("Going to YouTube Home.")
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.typewrite("https://www.youtube.com")
    pyautogui.press('enter')

def scroll_down():
    speak("Scrolling down.")
    pyautogui.scroll(-500)

def scroll_up():
    speak("Scrolling up.")
    pyautogui.scroll(500)

def increase_volume(amount=50):
    speak("Increasing volume.")
    for _ in range(amount // 10):
        pyautogui.press('volumeup')
        time.sleep(0.1)

def decrease_volume(amount=30):
    speak("Decreasing volume.")
    for _ in range(amount // 10):
        pyautogui.press('volumedown')
        time.sleep(0.1)

def seek_forward(seconds):
    speak(f"Forwarding {seconds} seconds.")
    for _ in range(seconds // 10):
        pyautogui.hotkey('shift', 'right')
        time.sleep(0.1)

def seek_backward(seconds):
    speak(f"Rewinding {seconds} seconds.")
    for _ in range(seconds // 10):
        pyautogui.hotkey('shift', 'left')
        time.sleep(0.1)

def play_video():
    if is_video_playing():
        speak("Playing the video.")
        pyautogui.press('space')
    else:
        speak("Opening the first video on screen.")
        pyautogui.moveTo(300, 400)
        pyautogui.click()
        time.sleep(3)

def extract_time(command):
    words = command.split()
    for word in words:
        if word.isdigit():
            return int(word)
    return 10

def process_command(command):
    if "play" in command:
        play_video()
    elif "pause video" in command or "pause" in command:
        speak("Pausing the video.")
        pyautogui.press('space')
    elif "scroll down" in command:
        scroll_down()
    elif "scroll up" in command:
        scroll_up()
    elif "search youtube for" in command:
        search_youtube(command.replace("search youtube for", "").strip())
    elif "volume up" in command:
        increase_volume()
    elif "volume down" in command:
        decrease_volume()
    elif "go forward" in command:
        time_to_seek = extract_time(command)
        seek_forward(time_to_seek)
    elif "go backward" in command or "rewind" in command:
        time_to_seek = extract_time(command)
        seek_backward(time_to_seek)
    elif "next video" in command:
        speak("Skipping to next video.")
        pyautogui.hotkey('shift', 'n')
    elif "previous video" in command:
        speak("Going back to previous video.")
        pyautogui.hotkey('shift', 'p')
    elif "open history" in command:
        open_history()
    elif "open subscriptions" in command:
        open_subscriptions()
    elif "go home" in command:
        open_home()
    elif "exit" in command or "quit" in command:
        speak("Goodbye!")
        exit()
    else:
        speak("Sorry, I didn't recognize that command.")

def main():
    open_youtube()
    while True:
        command = recognize_speech()
        if command:
            process_command(command)

if __name__ == "__main__":
    main()