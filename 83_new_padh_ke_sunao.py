import requests
import json
from win32com.client import Dispatch

def speak(str):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)

if __name__ == '__main__':
    speak("WELCOME TO THE SHIVAM GIRI NEWS CHANNEL, LETS MOVE TO THE TODAY'S TOP NEWS")
    speak("so, our first news is...")
    url = "https://newsapi.org/v2/top-headlines?country=in&apiKey=1ae5c19b047246a49e1a5f9125f58ac8"
    news = requests.get(url).text
    news_dict = json.loads(news)

    arts = news_dict['articles']
    for article in arts:
        speak(article['title'])
        print(article['title'])
        speak("moving on to the next news...")
    speak("thanks for listning, have a nice day")