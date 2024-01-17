import speech_recognition as sr
import win32com.client
import webbrowser
from openai import OpenAI
import os
import datetime
from config import apikey

speaker = win32com.client.Dispatch("SAPI.SpVoice")

chatStr = ""
def chat(query):
    client = OpenAI(api_key=apikey)
    global chatStr
    chatStr += f"Suman: {query}\n Roblo: "

    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {
                    "role": "user",
                    "content": chatStr
                }
            ],
            temperature=1,
            max_tokens=256,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )

        chatStr += f"{response.choices[0].message.content.strip()}\n"
        ans=response.choices[0].message.content.strip()
        print(ans)
        say(ans)

    except Exception as e:
        print("Error! Sorry API calls exhausted")
        say("Error! Sorry API calls exhausted")

def open_ai(prompt):
    client = OpenAI(api_key=apikey)
    text = f"OpenAI response for prompt: {prompt}\n ---------------------------------------------------\n\n"

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ],
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    text += response.choices[0].message.content.strip()
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    with open(f"D:\\ML\\VirtualAI\\Openai\\{prompt[0:20]}.txt", "w") as f:
        f.write(text)
    return (response.choices[0].message.content.strip())


def say(text):
    speaker.Speak(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)

    try:
        query = r.recognize_google(audio, language="en-in")
        print(f"Suman: {query}")
        return query
    except Exception as e:
        print("Sorry, I did not understand. Please try again.")
        say("Sorry, I did not understand. Please try again.")
        return "error"


def open_file(filename, ext):
    file_path = f"C:\\Users\\suman\\Documents\\{filename}"+"."+f"{ext}"
    try:
        os.system(f"start {file_path}")
    except Exception as e:
        print(f"Error opening file: {e}")

def play_song(filename):
    file_path = f"C:\\Users\\suman\\Music\\{filename}"+"."+"mp3"
    try:
        os.system(f"start {file_path}")
    except Exception as e:
        print(f"Error opening file: {e}")

def open_app(app_name):
    try:
        os.system(f"start {app_name}")
    except Exception as e:
        print(f"Error opening app: {e}")

if __name__ == '__main__':
    print("Welcome to Roblo")
    say("Welcome to Roblo")
    while True:
        print("Listening...")
        text = takeCommand()
        text = text.lower()

        if "error" in text:
            continue

        # OpenAI API Call
        elif "advanced" in text:
            print("-------------Advanced A.I Mode---------------")
            while True:
                print("Listening...")
                query = takeCommand()
                query = query.lower()
                if "exit" in query:
                    print("Exiting Advanced Mode....")
                    break
                prompt=open_ai(query)
                print(prompt)


        # Open Websites
        elif "open website" in text:
            website_name = text.replace("open website", "").strip()
            if website_name:
                url = f"https://{website_name.lower()}.com"
                say(f"Opening {website_name} website")
                webbrowser.open(url)
            else:
                say("I didn't catch the website name. Please try again.")

        # Open Files
        elif "open" in text and "file" in text:
            if "word" in text:
                filename=text.replace("open word file", "").strip()
                open_file(filename,"docx")
            elif "pdf" in text:
                filename=text.replace("open pdf file","").strip()
                open_file(filename,"pdf")

        # Play Songs
        elif "play" in text:
            filename = text.replace("play", "").strip()
            play_song(filename)

        # Tell Time
        elif "the time" in text:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"The time is {strfTime}")
            say(f"The time is {strfTime}")

        # Open Apps
        elif "open" in text:
            app_name = text.replace("open","").lstrip()
            print(app_name)
            if app_name:
                open_app(app_name)
                say(f"Opening {app_name}.")
            else:
                say("I didn't catch the application name. Please try again.")

        # Goodbye
        elif "bye" in text:
            print("Goodbye!...")
            say("Goodbye")
            break

        else:
            prompt = chat(text)


