import pyttsx3 as tts
import datetime as dt
import random
import os
import time
from selenium import webdriver
import pyautogui as p
import speech_recognition as sr
import smtplib as mail
from email.message import EmailMessage

from Credentials import *
from Emails import *

vibrations = tts.init()
vibrations.setProperty('rate', 200)


def tts_set_rate(rate):
    vibrations.setProperty('rate', rate)
    tts_speak_and_print('Voice rate has been set to', rate)


def tts_set_voice(number):
    voices_list = vibrations.getProperty('voices')
    if number == 0:
        vibrations.setProperty('voice', voices_list[0].id)
    elif number == 1:
        vibrations.setProperty('voice', voices_list[1].id)
    else:
        tts_speak_and_print('Beep beep. Invalid parameter. Use 1 for female voice and 0 for male voice.')
    tts_speak_and_print('Voice has been set to the type', number)


def tts_speak_and_print(text, value=None):  # 2 args sentence and value and 1 for speaking one line
    if value is not None:
        print(text, value)
        vibrations.say(text)
        vibrations.say(value)
    else:
        print(text)
        vibrations.say(text)
    vibrations.runAndWait()


# 1 for date. 2 for time and 3 for date and time
def date_and_time(option):
    current_date = dt.datetime.now().strftime("%A, %B %d, %Y")
    current_time = dt.datetime.now().strftime("%I: %M: %S")
    current_day_number_of_year = dt.datetime.now().strftime("%j")
    if option == '1':
        tts_speak_and_print("Today's date is ", current_date)
        tts_speak_and_print("and the day number is", current_day_number_of_year)
    elif option == '2':
        tts_speak_and_print("Current time is ", current_time)
    else:
        tts_speak_and_print("Today's date is ", current_date)
        tts_speak_and_print("and the day number is ", current_day_number_of_year)
        tts_speak_and_print("and the time is ", current_time)


def wish_you():
    current_hour = dt.datetime.now().hour
    if current_hour < 12:
        tts_speak_and_print('Good morning, Zen!')
    elif 12 <= current_hour < 17:
        tts_speak_and_print('Good afternoon, Zen.')
    elif 17 <= current_hour < 20:
        tts_speak_and_print('Good evening, Zen...')
    else:
        tts_speak_and_print('Have a wonderful night, Zen')


def get_comfy():
    wish_you()
    date_and_time(2)
    asking_to_help()


def asking_to_help():
    ask_list = ['How can I be of your assistance today?',
                'Hi, I"m vortex. Tell me what to do.',
                'What task can I make easier for you today?',
                'All set! Tell me what to do.',
                'A very nice day to make your life more easier, Zen. How can I help?',
                'Vortex at your service, sir.',
                "What's up, Zen? You know what to do."
                ]
    tts_speak_and_print(ask_list[random.randrange(0, 6)])


def fetch_attendance_from_UIMS():
    tts_speak_and_print('Tell the name of person.')
    uid_name = take_audio_command()
    uid = uids[uid_name]
    bro = webdriver.Edge(r'C:\Program Files\WebDriver\MicrosoftWebDriver.exe')
    bro.set_window_size(150, 150)
    bro.get('https://uims.cuchd.in/')

    username = bro.find_element_by_css_selector('#txtUserId')
    username.click()
    username.send_keys(uid)
    next_button = bro.find_element_by_css_selector('#btnNext')
    next_button.click()

    password = bro.find_element_by_css_selector('#txtLoginPassword')
    password.send_keys(uims_pass)
    password.click()
    login = bro.find_element_by_css_selector('#btnLogin')
    login.click()

    bro.get('https://uims.cuchd.in/UIMS/frmStudentCourseWiseAttendanceSummary.aspx?type=etgkYfqBdH1fSfc255iYGw==')

    time.sleep(3)
    table = bro.find_element_by_css_selector('#SortTable').text

    bro.close()

    now = dt.datetime.now().strftime("%A, %B %d, %Y --- %I: %M: %S")

#make a new txt file for storing scraped content from CUIMS
    f = open(r"D:\OneDrive - Chandigarh University\Scraped Content\Attendance_18BCS2450.txt", 'a')
    f.write('\n\n')
    f.write(now)
    f.write('\n\n')
    f.write(table)
    f.close()

    tts_speak_and_print('Attendance has been successfully scraped.\n')
    print(now, '\n\n', table)

    tts_speak_and_print('Opening file now.')
    os.startfile(r'D:\OneDrive - Chandigarh University\Scraped Content\Attendance_18BCS2450.txt')
    tts_speak_and_print('File has been opened')


def fetch_datasheet_from_UIMS():
    tts_speak_and_print('Tell the name of person.')
    uid_name = take_audio_command()
    uid = uids[uid_name]
    bro = webdriver.Edge(r'C:\Program Files\WebDriver\MicrosoftWebDriver.exe')
    bro.set_window_size(150, 150)

    bro.get('https://uims.cuchd.in/')

    username = bro.find_element_by_css_selector('#txtUserId')
    username.click()
    username.send_keys(uid)
    next_button = bro.find_element_by_css_selector('#btnNext')
    next_button.click()

    password = bro.find_element_by_css_selector('#txtLoginPassword')
    password.send_keys(uims_pass)
    password.click()
    login = bro.find_element_by_css_selector('#btnLogin')
    login.click()

    bro.get('https://uims.cuchd.in/UIMS/frmStudentDatesheet.aspx')

    time.sleep(3)
    datesheet_table = bro.find_element_by_css_selector(
        '#ContentPlaceHolder1_wucStudentDateSheet_gvStudentDateSheet > tbody').text

    bro.close()

    now = dt.datetime.now().strftime("%A, %B %d, %Y --- %I: %M: %S")

    f = open(r"D:\OneDrive - Chandigarh University\Scraped Content\Datesheet_18BCS2450.txt", 'a')
    f.write('\n\n')
    f.write(now)
    f.write('\n\n')
    f.write(datesheet_table)
    f.close()

    tts_speak_and_print('Date sheet has been successfully scraped.\n')
    print(now, '\n\n', datesheet_table)

    os.startfile(r'D:\OneDrive - Chandigarh University\Scraped Content\Attendance_18BCS2450.txt')


def run(command):
        p.hotkey('win', 'r')
        p.typewrite(str(command))
        p.press('enter')


def join_bb_class():
    sub, if_logged_in = take_text_command().split(' ')

    def login_procedure():
        run('https://cuchd.blackboard.com/')
        p.press('enter')
        time.sleep(9)
        # x, y = p.locateCenterOnScreen('sign.png', grayscale=True)
        p.click(920, 445)
        time.sleep(10)

    def join_class_page():
        p.moveTo(110, 460, 15)
        p.click()
        p.click(180, 527)

    run('ed')
    time.sleep(3)

    for _ in range(0, 3):
        p.hotkey('win', 'up')

    if int(if_logged_in) < 1:
        login_procedure()

    if sub == 'apti':
        run('https://cuchd.blackboard.com/ultra/courses/_21859_1/outline')
    elif sub == 'apti_w':
        run('https://cuchd.blackboard.com/ultra/courses/_22766_1/outline')
    elif sub == 'is':
        run('https://cuchd.blackboard.com/ultra/courses/_19154_1/outline')
    elif sub == 'isl':
        run('https://cuchd.blackboard.com/ultra/courses/_19181_1/outline')
    elif sub == 'mad':
        run('https://cuchd.blackboard.com/ultra/courses/_20152_1/outline')
    elif sub == 'nos':
        run('https://cuchd.blackboard.com/ultra/courses/_20182_1/outline')
    elif sub == 'project':
        run('https://cuchd.blackboard.com/ultra/courses/_19888_1/outline')
    elif sub == 'pblj':
        run('https://cuchd.blackboard.com/ultra/courses/_20167_1/outline')
    elif sub == 'pbljl':
        run('https://cuchd.blackboard.com/ultra/courses/_19710_1/outline')
    elif sub == 'ss':
        run('https://cuchd.blackboard.com/ultra/courses/_21730_1/outline')
    elif sub == 'ss_w':
        run('https://cuchd.blackboard.com/ultra/courses/_22815_1/outline')
    elif sub == 'tt':
        run('https://cuchd.blackboard.com/ultra/courses/_19666_1/outline')
    elif sub == 'toc':
        run('https://cuchd.blackboard.com/ultra/courses/_20136_1/outline')
    else:
        p.moveTo(530, 335)
        p.moveTo(1570, 340, 20)
        run('https://cuchd.blackboard.com/ultra/course')

    join_class_page()


def take_audio_command():
    my_voice = sr.Recognizer()
    flag = 1
    while flag is not None:
        with sr.Microphone() as source:
            print('Listening...')
            my_voice.pause_threshold = 1
            audio = my_voice.listen(source)
        try:
            print('Recognizing')
            query = my_voice.recognize_google(audio, language='en-IN')
            print(query)
            flag = None
        except Exception as e:
            print(e)
            tts_speak_and_print("Please repeat what you've just said.")
    return query


def take_text_command():
    tts_speak_and_print('Type your query.')
    return input()


def send_mail():
    server = mail.SMTP('smtp.gmail.com', 587)
    server.starttls()
    print('Logging in')
    server.login(username, password)
    print('Logged in successfully')
    email = EmailMessage()
    stop = 1
    email['From'] = username
    tts_speak_and_print('To whom you want to send an e-mail?')
    while stop is not None:
        name_of_receiver = take_audio_command()
        receiver = email_list[name_of_receiver]
        tts_speak_and_print('The corresponding e-mail id is: ', receiver)
        tts_speak_and_print('Is this the correct e-mail id?')
        check = take_audio_command()
        if check == 'yes':
            stop = None
        elif check == 'no':
            tts_speak_and_print('Lets try again with different name then...')
    email['To'] = receiver

    stop = 1

    while stop is not None:
        tts_speak_and_print('What should be the subject of the e-mail?')
        subject = take_audio_command()
        tts_speak_and_print('Is this the subject you want to add?\n-------------------------------------')
        print(subject)
        print('-------------------------------------\n')
        check = take_audio_command()
        if check == 'yes':
            stop = None
        elif check == 'no':
            tts_speak_and_print('Please repeat the subject again then.')
    email['Subject'] = subject

    stop = 1

    tts_speak_and_print('Say the content to be added.')
    content = take_audio_command()
    while stop is not None:
        tts_speak_and_print('Content entered is:\n')
        print(content)
        tts_speak_and_print('\nIs this is the right content?')
        check = take_audio_command()
        if check == 'yes':
            stop = None
        elif check == 'no':
            tts_speak_and_print('Okay, say the content of the e-mail again.')
    email.set_content(content)
    tts_speak_and_print('Email which will be sent is as follows:')
    print('Subject: ', subject, end='\n\n')
    print('Body: ', content)
    server.send_message(email)
    tts_speak_and_print('Email has been sent to ', email_list[name_of_receiver])
    server.close()


def whatsapp_message():
    tts_speak_and_print('Type the number to send WhatsApp message.')
    number = take_text_command()
    flag = 1
    while flag is not None:
        tts_speak_and_print('Speak the message you want to send.')
        message = take_audio_command()
        tts_speak_and_print('Is this the correct message?\n---------------------------')
        print(message+'---------------------------\n')
        tts_speak_and_print('Please confirm.')
        afferm = take_audio_command()
        if afferm == 'yes':
            flag = None
        else:
            tts_speak_and_print('Okay, speak the message again')

    url = 'https://wa.me/91'+number+'?text='+message
    tts_speak_and_print('Sending message.')
    run(url)
    p.sleep(3)
    p.press('enter')
    tts_speak_and_print('Your message has been sent.')


def search(sentence):
    run('https://www.google.com/search?q='+sentence)


def play(song):
    run('https://open.spotify.com/search/'+song)


def movie(name):
    run('https://m.imdb.com/find?q='+name)


def wiki(key):
    run('https://en.wikipedia.org/wiki/Special:Search?search='+key)


if __name__ == "__main__":
    #get_comfy()
    while True:
        query = take_audio_command().lower()
        if 'date and time' in query:
            date_and_time(3)
        elif 'mail' in query:
            send_mail()
        elif 'attendance' in query:
            fetch_attendance_from_UIMS()
        elif 'date sheet' in query:
            fetch_datasheet_from_UIMS()
        elif 'join' in query:
            join_bb_class()
        elif 'message' in query:
            whatsapp_message()
        elif 'search' in query:
            query = query.replace('search', '')
            search(query)
        elif 'play' in query:
            query = query.replace('play', '')
            play(query)
        elif 'movie' in query:
            query = query.replace('play', '')
            movie(query)
        elif 'wikipedia' in query:
            query = query.replace('wikipedia', '')
            wiki(query)
        elif 'stop' in query:
            tts_speak_and_print('Exiting now.')
            break
