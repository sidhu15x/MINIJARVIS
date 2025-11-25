import pyttsx3  #pip install pyttsx3
import pyttsx3.driver #pip install pyttsx3
import datetime
import speech_recognition as sr
import webbrowser
import os
from translate import Translator
import keyboard
import win32clipboard
import clipboard
import requests
import wmi
import psutil
import smtplib
import wikiquote
import ctypes
import pyautogui
from time import sleep
from tkinter import *
import wikipedia
import random
import pygame
import wolframalpha
from PyDictionary import PyDictionary

engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
engine.setProperty('voice',voices[0].id)
engine.setProperty('rate', 180)

def game_snake():
    pygame.init()
    clock = pygame.time.Clock()
    orange = (255, 123, 7)
    black = (0, 0, 0)
    red = (213, 0, 0)
    green = (0, 255, 0)
    blue = (0, 0, 255)

    dis = pygame.display.set_mode((600, 600))
    pygame.display.set_caption('Snake Game')
    snake_block = 10
    snake_speed = 15
    snake_list = []
    int()

    def snake(snake_block, snake_list):
        for x in snake_list:
            pygame.draw.rect(dis, orange, (int(x[0]), int(x[1]), snake_block, snake_block))

    def snake_game():
        game_over = False
        game_end = False
        x1 = 600 / 2
        y1 = 600 / 2
        x1_change = 0
        y1_change = 0
        foodx = round(random.randrange(0, 600 - snake_block) / 10.0) * 10.0
        foody = round(random.randrange(0, 600 - snake_block) / 10.0) * 10.0
        snake_list = []
        length_of_snake = 1
        while not game_over:
            while game_end == True:
                game_end = False
                score = length_of_snake - 1
                pygame.quit()
                speak('oops, you lost')
                speak('Your Score was {}'.format(score))
                start()


            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    game_over = True
                if event.type == pygame.KEYDOWN:
                    if event.key == pygame.K_LEFT:
                        x1_change = -snake_block
                        y1_change = 0
                    if event.key == pygame.K_RIGHT:
                        x1_change = snake_block
                        y1_change = 0
                    if event.key == pygame.K_UP:
                        y1_change = -snake_block
                        x1_change = 0
                    if event.key == pygame.K_DOWN:
                        y1_change = snake_block
                        x1_change = 0
            if x1 >= 600 or x1 < 0 or y1 >= 600 or y1 < 0:
                game_end = False
            x1 += x1_change
            y1 += y1_change
            dis.fill(black)
            pygame.draw.rect(dis, green, [int(foodx), int(foody), int(snake_block), int(snake_block)])
            snake_head = []
            snake_head.append(x1)
            snake_head.append(y1)
            snake_list.append(snake_head)
            if len(snake_list) > length_of_snake:
                del snake_list[0]

            for x in snake_list[:-1]:
                if x == snake_head:
                    game_end = True
            snake(snake_block, snake_list)
            pygame.display.update()

            if (x1 == 600 or x1 == 0) or (y1 == 600 or y1 == 0):
                game_end = True
            if x1 == foodx and y1 == foody:
                foodx = round(random.randrange(0, 600 - snake_block) / 10.0) * 10.0
                foody = round(random.randrange(0, 600 - snake_block) / 10.0) * 10.0
                length_of_snake += 1
            clock.tick(snake_speed)
        pygame.quit()
        quit()

    snake_game()

def secs2hours(secs):
    seconds = secs % (24 * 3600)
    hour = seconds // 3600
    seconds %= 3600
    minutes = seconds // 60
    seconds %= 60

    return "{}hours {}minutes and {}seconds".format(hour, minutes, seconds)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def takecommand():
    query=""
    r=sr.Recognizer()
    r.energy_threshold = 400
    with sr.Microphone() as source:
        print("Listening Your Orders....")
        audio=r.listen(source)

        try:
            print("Recognizing...")
            query=r.recognize_google(audio, language='en-in')
            print("User Said:",query,'\n')

        except Exception as e:
                print("Say that again please..")
                query=takecommand()
                return query
    return query

def num2words(num):
    under_20 = ['Zero','One','Two','Three','Four','Five','Six','Seven','Eight','Nine','Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen']
    tens = ['Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety']
    above_100 = {100: 'Hundred',1000:'Thousand', 100000:'Lakhs', 10000000:'Crores'}

    if num < 20:
         return under_20[(int)(num)]

    if num < 100:
        return tens[(int)(num/10)-2] + ('' if num%10==0 else ' ' + under_20[(int)(num%10)])

    # find the appropriate pivot - 'Million' in 3,603,550, or 'Thousand' in 603,550
    pivot = max([key for key in above_100.keys() if key <= num])

    return num2words((int)(num/pivot)) + ' ' + above_100[pivot] + ('' if num%pivot==0 else ' ' + num2words(num%pivot))

def wishMe():
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak('Good Morning')
    elif hour>=12 and hour<18:
        speak("Good afternoon")
    else:
        speak('good evening')

    speak('i am jarvis. Sir, Please tell me how may i help you')
    speak('You can activate me by tapping the microphone')

def sendEmail(reciever,subject,content):
    server=smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('gursimran1822000@gmail.com','08310831')
    content='Subject: {}\n\n{}'.format(subject, content)
    server.sendmail('gursimran1822000@gmail.com',reciever,content)
    server.close()

def functionality(query):
    if 'search wikipedia for' in query:
        speak('searching...')
        query = query.replace('search wikipedia for', '')
        results = wikipedia.summary(query, sentences=2)
        speak("According to wikipedia")
        print(results)
        speak(results)


    elif 'what is antonym of' in query:
        replies = ['ok, let me check', 'ok, let me find', 'one moment please', 'one second please', 'just a moment',
                   'just a second', 'just a minute']
        speak(random.choice(replies))
        query = query.replace("what is antonym of", "")
        dictionary = PyDictionary()
        antonym_list = dictionary.antonym(query)
        speak('antonyms for {} are'.format(query))
        for i in range(len(antonym_list)):
            speak(antonym_list[i])

    elif 'motivate me' in query:
        list = wikiquote.quotes('Motivation')
        engine.setProperty('rate', 150)
        speak(random.choice(list))
        engine.setProperty('rate', 180)

    
    elif 'what is synonym of' in query:
        replies = ['ok, let me check', 'ok, let me find', 'one moment please', 'one second please', 'just a moment',
                   'just a second', 'just a minute']
        speak(random.choice(replies))
        query = query.replace("what is synonym of", "")
        dictionary = PyDictionary()
        synonims_list = dictionary.synonym(query)
        speak('synonyms for {} are'.format(query))
        for i in range(len(synonims_list)):
            speak(synonims_list[i])


    elif 'question' in query:
        replies_1=['what is it, tell me, i will try to answer','let me clear it, what is it,sir','i am here to help you, tell me']
        speak(random.choice(replies_1))
        question=takecommand().lower()
        replies=['ok, let me check','ok, let me find','one moment please','one second please','just a moment','just a second','just a minute']
        speak(random.choice(replies))
        try:
            app_id='985GUJ-QR58PPJLJT'
            client=wolframalpha.Client(app_id)
            res=client.query(question)
            answer=next(res.results).text
            speak(answer)
        except:
            speak('sorry, i dont know the answer')
            speak('google is much smarter that me')
            speak('do you want me to search it there')
            choice=takecommand().lower()
            if choice=='yes':
                webbrowser.get('windows-default').open('https://google.com/search?q={}'.format(question))



    elif 'open youtube' in query:
        speak("Opening Youtube for you, Sir")
        webbrowser.get('windows-default').open('https://www.youtube.com')


    elif 'snake game' in query:
        speak('ok, enjoy')
        game_snake()

    elif 'date' in query:
        strdate = datetime.datetime.today().strftime("%d %m %y")
        speak("Sir today's date is %s" % strdate)

    elif 'screenshot' in query:
        speak('Will you please provide a filename for it')
        filename=takecommand().lower()
        myScreenshot = pyautogui.screenshot()
        myScreenshot.save(r'C:\\Users\\Public\\{}.png'.format(filename))
        speak('screenshot taken sir, check your desktop, i saved it there for you')
        speak("do you want me to show it")
        choice=takecommand().lower()
        if choice=='yes':
            speak('just a second, sir')
            os.startfile("C:\\Users\\Public\\{}.png".format(filename))
            sleep(1)
            speak("looks good")

    elif 'search youtube for' in query:
        speak('searching...')
        query = query.replace('search youtube for', '')
        url='https://www.youtube.com/results?search_query={}'.format(query)
        webbrowser.get('windows-default').open(url)


    elif 'brightness' in query:
        if 'decrease brightness' in query:
            speak("decreasing brightness")
            dec = wmi.WMI(namespace='wmi')
            methods = dec.WmiMonitorBrightnessMethods()[0]
            methods.WmiSetBrightness(30,0)
        elif 'increase brightness' in query:
            speak("increasing brightness")
            ins = wmi.WMI(namespace='wmi')
            methods = ins.WmiMonitorBrightnessMethods()[0]
            methods.WmiSetBrightness(100, 0)

    elif 'battery status' in query:
        battery = psutil.sensors_battery()
        plugged = battery.power_plugged
        percent = int(battery.percent)
        time_left = secs2hours(battery.secsleft)
        if plugged==True:
            speak("Don't worry sir, charger is plugged in")

        elif percent > 40:
            speak('sir, currently there is no need to charge me, i can survive for' + time_left)

        else:
            speak('sir, please connect the charger as soon as possible, i can survive only for' + time_left)


    elif 'lock my' in query:
        speak('ok, sir')
        ctypes.windll.user32.LockWorkStation()


    elif 'translate' in query:
        speak('What do you want me to translate')
        content = takecommand().lower()
        speak("In which language do you want to translate, sir")
        language=takecommand().lower()
        translator = Translator(to_lang=language)
        translation = translator.translate(content)
        clipboard.copy(translation)
        speak(translation)
        speak("I copied the translation to clipboard, for your convenience")


    elif 'news' in query:
        pass

    elif 'thank you' in query:
        speak("Welcome Sir")

    elif 'open google' in query:
        speak("Do you want me to search for you sir?")
        choice=takecommand().lower()
        if choice=='yes':
            speak("what do you want me to search sir?")
            search_item=takecommand().lower()
            speak("Searching for")
            speak(search_item)
            webbrowser.get('windows-default').open('https://google.com/search?q={}'.format(search_item))
        else:
            speak("Opening Google for you, Sir")
            webbrowser.get('windows-default').open('http://www.google.com')

    elif 'joke' in query:
        url = "https://joke3.p.rapidapi.com/v1/joke"

        headers = {
            'x-rapidapi-host': "joke3.p.rapidapi.com",
            'x-rapidapi-key': "7d742bae65mshf3fc2bd05d29eb8p171eebjsn49dd0d6ffa1c"
        }

        response = requests.request("GET", url, headers=headers)
        dictionary = response.json()
        engine.setProperty('rate', 150)
        speak(dictionary['content'])


    elif 'open calculator' in query:
        speak('ok')
        os.system('calc')

    elif 'open task manager' in query:
        speak('ok')
        os.system('taskmgr.exe')

    elif 'open this pc' in query:
        speak('ok')
        os.system('explorer.exe')


    elif 'locate' in query:
        url='https://www.google.nl/maps/place/'
        query=query.replace('locate','')
        speak('Locating {}, Sir'.format(query))
        url=url+query+'/&amp;'
        webbrowser.get('windows-default').open(url)
        sleep(1)
        speak("here is the location of{}".format(query))

    elif 'read text' in query:
        speak("Ok sir, Copy any text to the clipboard, and i will read that text for you")
        keyboard.wait('ctrl+c')
        sleep(0.5)
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        speak(data)

    elif 'open stackoverflow' in query:
        speak("Opening StackOverflow for you, Sir")
        webbrowser.open('stackoverflow.com')

    elif 'play music' in query:
        speak("Playing your favorite Song")
        speak('enjoy')
        music_dir = 'D:\\Songs'
        songs = os.listdir(music_dir)
        print(songs)
        os.startfile(os.path.join(music_dir, songs[0]))


    elif 'the time' in query:
        strtime = datetime.datetime.now().strftime("%H:%M:%S")
        speak("Sir the time is")
        speak(strtime)

    elif 'old are you' in query:
        speak("I am a little baby sir")

    elif 'your name' in query:
        speak('myself Jarvis, sir')

    elif 'who created you' in query:
        speak('My Creator is Gursimranjeet Singh')

    elif 'hello' and 'everyone' in query:
        speak('Hello Everyone! My self Jarvis')
        speak("I hope you all are doing great")

    elif 'hello' in query:
        speak("Hello, Welcome Sir")

    elif 'technology' in query:
        speak("I am created in Python programming language")

    elif 'change your gender' in query:
        speak("Why do you want me to change my gender")
        sleep(0.5)
        speak('Its not an easy task but still i will do it for you')
        gender=engine.getProperty('voice')
        if gender=='HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_DAVID_11.0':
            engine.setProperty('voice',voices[1].id)
        else:
            engine.setProperty('voice', voices[0].id)
        sleep(1)
        speak("Affermative")
        print(gender)

    elif 'open notepad' in query:
        speak('Opening notepad')
        os.system('notepad')

    elif 'close' in query:
        keyboard.press_and_release('alt+f4')

    elif 'type for me' in query:
        speak('What do you want me to type')
        input_text=takecommand()
        keyboard.write(input_text)
        speak("typing done sir")

    elif 'weather' in query:
        key='4c32e944ea57435751b017775f1e2cf6'
        res=requests.get('http://api.openweathermap.org/data/2.5/weather?q={}&appid={}'.format('patiala',key))
        data=res.json()
        temperature=data['main']['temp']
        temperature=temperature-273
        wind_speed=data['wind']['speed']
        description=data['weather'][0]['description']
        speak("In Patiala")
        speak("Its {} outside".format(description))
        speak("The temperature is {} degree celcius".format(int(temperature)))
        speak("The wind is flowing at a speed of {} meter per second".format(wind_speed))

    elif "your capabilities" in query or "what can you do for me" in query:
        speak("I am having different capabilities")
        speak("I can search in wikipedia for you")
        speak("I can open youtube, google, stackoverflow")
        speak("I can play music for you, I can tell time")
        speak("I can also send emails for you")
        speak("I can open chrome and others softwares too, for you")
        speak("I will be happy to help you")
        speak("I am designed to talk formally with users, and very soon i will be having learning abilities")
        speak("well, My creator Gursimranjeet singh has made me much smarter than other virtual assistant's")
        speak("currently, i am under development, and will have more features in future")
        speak("Thank you very much for asking")

    elif "select all" in query:
        speak('selecting all')
        keyboard.press_and_release('ctrl+a')

    elif "copy" in query:
        speak("copying")
        keyboard.press_and_release('ctrl+c')
        speak('done')

    elif "undo" in query:
        speak("ok")
        keyboard.press_and_release('ctrl+z')

    elif "paste" in query:
        speak('pasting')
        keyboard.press_and_release('ctrl+v')

    elif "who are you" in query or "tell me about yourself" in query or "introduce yourself" in query:
        speak("I am baby jarvis, speed 1 tera hertz, memory 1 terabyte")
        speak("I am a creation of Harmanjeet Singh")
        speak("I am currently under development and i am learning many more things")
    elif 'open chrome' in query:
        speak("Opening Google Chrome")
        chrome_path = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
        os.startfile(chrome_path)

    elif 'open pycharm' in query:
        pycharm_location = "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2020.1.1\\bin\\pycharm64.exe"
        os.startfile(pycharm_location)


    elif 'bye' in query:
        speak("Good Bye,Sir")
        root.destroy()
        exit()

    elif 'email to me' in query:
        try:
            speak("What's the subject of the email")
            subject=takecommand()
            speak("What should i type in the mail")
            content = takecommand()
            reciever = "gs1822000@gmail.com"
            sendEmail(reciever,subject, content)
            speak("email has been sent")
        except Exception as e:
            print(e)
            print("Sorry, I am not able to send email at this moment")
    else:
        speak("Sorry, i am not able to understand you")

def start():
    while True :
        query = takecommand().lower()
        functionality(query)

def buttonreleased(event):
    speak("Ready and listening sir")
    start()

wishMe()
root = Tk()
root.configure(bg='white')
root.title("Jarvis Mic")
root.lift()
root.iconbitmap('robot.ico')
photo = PhotoImage(file='microphone (1).png')
label = Label(root, image=photo, bg="white")
label.grid()
label.bind("<ButtonPress>", buttonreleased)
mainloop()
