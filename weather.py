import speech_recognition as sr
import pyttsx3
import requests

# Initialize text-to-speech engine
engine = pyttsx3.init()

# OpenWeatherMap API key
API_KEY = "da06cb560539c44265f3b751c6d9f489"  # Your working key

def speak(text):
    """Speaks the given text."""
    print(f"Nexa: {text}")
    engine.say(text)
    engine.runAndWait()

def listen():
    """Listens and returns recognized speech as text."""
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=1)
        print("Listening...")
        audio = recognizer.listen(source)
        try:
            command = recognizer.recognize_google(audio)
            print(f"✅ You said: {command}")
            return command.lower().strip()
        except sr.UnknownValueError:
            speak("Sorry, I didn't catch that. Please repeat.")
            return ""
        except sr.RequestError as e:
            speak("Speech service is unavailable.")
            print(f"RequestError: {e}")
            return ""

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

def nexa():
    """Runs the voice assistant."""
    speak("Hello! Say the name of a city to get the weather report. Say stop or bye to exit.")

    while True:
        speak("Please say a city name.")
        city = listen()

        if not city:
            continue

        # ✅ Debug line to show what was recognized
        print(f"DEBUG: Checking for exit in: '{city}'")

        if "stop" in city or "bye" in city:
            speak("Goodbye! Have a great day.")
            break

        weather_report = get_weather(city)
        speak(weather_report)


if __name__ == "__main__":
    try:
        nexa()
    except KeyboardInterrupt:
        speak("Exiting. Goodbye!")