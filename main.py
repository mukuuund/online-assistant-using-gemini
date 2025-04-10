import speech_recognition as sr
import json
import os
import win32com.client
import webbrowser
import datetime
import google.generativeai as genai
import smtplib
from email.message import EmailMessage
import re

from dotenv import load_dotenv
import os

load_dotenv()

genai.configure(api_key=os.getenv("API_KEY"))


model = genai.GenerativeModel('gemini-1.5-flash-latest')


chat_file="chat_history.json"

if os.path.exists(chat_file):
    with open(chat_file,"r",encoding="utf-8") as f:
        chat_history=json.load(f)
else:
     chat_history = []

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(text):
    speaker.Speak(text)


def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio,language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Sorry! Some error occurred"


from datetime import datetime

def log_chat_json(user_query,jarvis_response):
    entry={
        "timestamp":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "user": user_query.strip(),
        "jarvis":jarvis_response.strip()
    }
    chat_history.append(entry)

    with open(chat_file, "w", encoding="utf-8") as f:
        json.dump(chat_history, f, indent=4)

def clean_email_input(raw_input):
    email = raw_input.lower().strip()
    email = email.replace(" at ", "@").replace(" dot ", ".")
    email = email.replace(" ", "")  # Remove any leftover spaces
    return email

def extract_url(txt):
    urls=re.findall(r'https?://\S+',txt)
    if urls:
        url = urls[0].strip('"\',.;)`\n')
        return url
    else:
        return None


def send_email(to, about):
    try:
        email=EmailMessage()
        email['From']='xyzmail'
        q1=f"Write the subject of an email to {to} on {about}"
        resp = model.generate_content(q1)
        subject=resp.text
        q2=f"Write a body of an email to {to} on {subject}"
        resp = model.generate_content(q2)
        body=resp.text
        email['To'] = to.strip().replace('\n', '').replace('\r', '')
        email['Subject'] = subject.strip().replace('\n', '').replace('\r', '')
        email.set_content(body.strip())

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login("YOURMAIL@gmail.com", os.getenv("EMAIL_PASS"))
            smtp.send_message(email)

        say("Email sent successfully.")
        print("Email sent to ",to)

    except Exception as e:
        say("Sorry! I was not able to send the email")
        print("Email error", e)



if __name__ == '__main__':
    print("Pycharm")
    say("Hello I am Jarvis A I")
    while True:
        print("Listening...")
        query = takeCommand()
        if query is None or "error" in query.lower():
            say("Sorry, I didn't catch that. Can you repeat?")
            continue

        if "the time" in  query:
            strftime=datetime.now().strftime("%H:%M:%S")
            say(f"The time is {strftime}")

        elif query.lower()=="jarvis stop" or query.lower()=="exit":
            say("Goodbye! Have a great day.")
            break;

        elif "send email" in query.lower():
            say("Who should I send the email to?")
            print("listening....")
            recipient = takeCommand()
            to = clean_email_input(recipient)
            say("What is the email about ?")
            print("listening....")
            about= takeCommand()
            send_email(to,about)

        else:
            try:
                if "open" in  query.lower():
                    response = model.generate_content(f"just give me the complete url for the given query {query} for opening it via python webbrowser module for my project")
                    url=extract_url(response.text)
                    if url:
                        say("Opening requested website")
                        print(url)
                        webbrowser.open(url)
                    else:
                        print(url)
                        say("Sorry, I couldn't find a valid URL.")
                else:
                    response = model.generate_content(query)
                    result = response.text
                    say(result)
                    print(result)
                    log_chat_json(query, result)
            except Exception as e:
                say("Sorry, I couldn't process that.")
                print("Gemini API error:", e)