import time
import datetime
import threading
import tkinter as tk
import speech_recognition as sr
import pyttsx3
import sys

# Initialize speech engine
engine = pyttsx3.init()

def talk_back(text):
    print(text)
    engine.say(text)
    engine.runAndWait()

def listen_in_background():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()
    stop_listening = recognizer.listen_in_background(mic, callback)
    return stop_listening

def callback(recognizer, audio):
    global stopwatch_running
    try:
        command = recognizer.recognize_google(audio).lower()
        print(f"Recognized command: {command}")

        if "start" in command and not stopwatch_running:
            root.after(0, start_stopwatch)

        elif "stop" in command and stopwatch_running:
            root.after(0, stop_stopwatch)

        elif "close stopwatch" in command or "exit stopwatch" in command:
            talk_back("Closing the stopwatch.")
            root.quit()

    except sr.UnknownValueError:
        pass
    except sr.RequestError:
        talk_back("Could not request results, check your internet.")

# GUI Update Functions
def update_timer():
    if stopwatch_running:
        elapsed = time.time() - start_time
        timer_var.set(f"{elapsed:.2f} seconds")
        root.after(100, update_timer)

# Stopwatch control
def start_stopwatch():
    global stopwatch_running, start_time
    if not stopwatch_running:
        stopwatch_running = True
        start_time = time.time()
        start_label.config(text=f"Started at: {datetime.datetime.now().strftime('%H:%M:%S')}")
        talk_back("Stopwatch started.")
        update_timer()

def stop_stopwatch():
    global stopwatch_running
    if stopwatch_running:
        stopwatch_running = False
        end_time = time.time()
        elapsed = end_time - start_time
        stop_label.config(text=f"Stopped at: {datetime.datetime.now().strftime('%H:%M:%S')}")
        timer_var.set(f"Total: {elapsed:.2f} seconds")
        talk_back(f"Stopwatch stopped. Total time: {elapsed:.2f} seconds.")

# GUI Setup
root = tk.Tk()
root.title("Voice Controlled Stopwatch")
root.geometry("320x220")

timer_var = tk.StringVar(value="0.00 seconds")
stopwatch_running = False
start_time = 0

tk.Label(root, text="Stopwatch Timer", font=("Arial", 14)).pack(pady=10)
tk.Label(root, textvariable=timer_var, font=("Arial", 18)).pack()

start_label = tk.Label(root, text="")
start_label.pack()
stop_label = tk.Label(root, text="")
stop_label.pack()

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

tk.Button(btn_frame, text="Start", command=start_stopwatch).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Stop", command=stop_stopwatch).pack(side=tk.LEFT, padx=5)

# Start voice recognition in background
stop_listening = listen_in_background()

# Close listener and window cleanly
def on_closing():
    stop_listening(wait_for_stop=False)
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()