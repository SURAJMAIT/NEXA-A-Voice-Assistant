import speech_recognition as sr
import pyttsx3
from nltk.corpus import wordnet
import nltk

# Download wordnet dataset
nltk.download('wordnet')

# Initialize text-to-speech engine
engine = pyttsx3.init()

# Function to get word meaning and speak it
def get_meaning(word):
    synsets = wordnet.synsets(word)
    if synsets:
        definition = synsets[0].definition()
        response = f"The meaning of {word} is: {definition}"
    else:
        response = f"Sorry, I couldn't find the meaning of {word}."
    
    print(response)
    engine.say(response)
    engine.runAndWait()

# Function to capture voice input in a loop
def listen_and_define():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Speak a word to get its meaning.\n Say 'Stop' to exit.")
        recognizer.adjust_for_ambient_noise(source)

        while True:
            try:
                print("\nSpeak a word to get its meaning.")
                audio = recognizer.listen(source)
                word = recognizer.recognize_google(audio).lower()
                print(f"You said: {word}")

                # Stop the loop if user says "stop"
                if word == "stop":
                    print("Okay, stopping the program now.")
                    engine.say("Okay, stopping the program now.")
                    engine.runAndWait()
                    break
                
                get_meaning(word)

            except sr.UnknownValueError:
                print("Sorry, I couldn't understand. Please try again.")
            except sr.RequestError:
                print("Could not request results, check your connection.")

if __name__ == "__main__":
    listen_and_define()
