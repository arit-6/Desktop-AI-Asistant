import speech_recognition as sr
import openai
import webbrowser
import win32com.client
from config import apikey
import os


def ai(prompt):
    openai.api_key = apikey

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    value = response["choices"][0]["text"]
    print(value)
    return value

def say(text):
    spkr = win32com.client.Dispatch("SAPI.SpVoice")
    spkr.Speak(text)


def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing......")
            qs = r.recognize_google(audio, language="en-in")
            print(f"User said: {qs}")
            return qs
        except Exception as e:
            return "Some error occurred,  Sorry from JARVIS"


print("STROKE A.I")
say("STROKE A I")


while True:

    print("Listening......")
    say("Listening")
    query = takecommand()

    if "open" in query.lower():
        words = query.split()
        print(words[1])
        try:
            webbrowser.open(f"https://{words[1].lower()}.com")
        except Exception:
            say("Some error occurred,  Sorry from STROKE")
        say(f"Opening {words[1]}")

    elif "write" in query.lower():
        say("writing")
        if not os.path.exists("Letters"):
            os.mkdir("Letters")
        with open(f"Letters/{''.join(query.split('write me a ')[1:])}.txt", "w") as file:
            resfrmopenai = ai(query)
            file.write(resfrmopenai)

    elif "quit" in query.lower():
        exit()

    elif "nothing" in query.lower():
        say("Have a great day...")
        exit()

    elif "stop" in query.lower():
        exit()

    elif "close" in query.lower():
        exit()

    else:
        resfrmopenai = ai(query)
        say(resfrmopenai[:200])
