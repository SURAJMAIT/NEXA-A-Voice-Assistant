from utils.speech_utils import speak, listen
from modules import (
    app_launcher,
    dictionary,
    reminder,
    stopwatch,
    notepad_writer,
    web_navigator,
    weather,
    youtube_controller
)

def main():
    speak("Hey, I am Nexa. How can I help you?")
    while True:
        command = listen()
        if not command:
            continue

        if "open" in command and "app" not in command:
            app_name = command.replace("open", "").strip()
            app_launcher.open_app(app_name)

        elif "close" in command and "app" not in command:
            app_name = command.replace("close", "").strip()
            app_launcher.close_app(app_name)

        elif "define" in command:
            word = command.replace("define", "").strip()
            dictionary.get_meaning(word)

        elif "reminder" in command or "set reminder" in command:
            reminder.set_reminder()

        elif "stopwatch" in command:
            stopwatch.run_stopwatch()

        elif "notepad" in command or "write" in command or "word" in command or "notes" in command:
            notepad_writer.main()

        elif "search" in command or "browser" in command:
            web_navigator.execute(command)

        elif "weather" in command:
            weather.main()

        elif "youtube" in command or "play video" in command:
            youtube_controller.control(command)

        elif "exit" in command or "bye" in command:
            speak("Goodbye! Have a great day.")
            break

        else:
            speak("Sorry, I didn't understand. Please try again.")

if __name__ == "__main__":
    main()