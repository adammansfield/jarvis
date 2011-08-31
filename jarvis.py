import win32com.client
import pywapi
import random
import string
import time
import shutil
import imaplib
import re
import os
import sys

sys.path.append(os.getcwd() + "\\config")


# Default User Config 
first_name = 'John'
last_name = 'Doe'
city = 'New York'
location_id = '10001'
email_username = '1234@gmail.com'
email_password = '1234'

# Import custom userconfig
try:
    from userconfig import *
except ImportError:
    shutil.copyfile(os.getcwd() + "\\config\\default.py", os.getcwd() + "\\config\\userconfig.py")


def get_greeting():
    option = random.uniform(0,100)
    if   (option<=25): greeting_name = ', Mr. ' + last_name + '. '
    elif (option<=50): greeting_name = ' ' + first_name + '. '
    elif (option<=75): greeting_name = ' sir. '
    else:              greeting_name = '. '

    hour = int(time.strftime("%H", time.localtime()))
    option = random.uniform(0,100)
    if   (option<=20):              greeting = 'Hello'          + greeting_name
    elif (option<=40):              greeting = 'Hi'             + greeting_name
    elif (hour >=0 and hour < 5):   greeting = 'Good night'     + greeting_name
    elif (hour >=5 and hour < 12):  greeting = 'Good morning'   + greeting_name
    elif (hour >=12 and hour < 17): greeting = 'Good afternoon' + greeting_name
    else:                           greeting = 'Good evening'   + greeting_name
    return greeting

def get_signoff():
    option = random.uniform(0,100)
    if   (option<=33): signoff_name = ' ' + first_name + '. '
    elif (option<=66): signoff_name = ' sir. '
    else:              signoff_name = '. '

    hour = int(time.strftime("%H", time.localtime()))
    if   (hour >=0  and hour < 5):  signoff = 'That is all for tonight. Have a pleasant sleep'          + signoff_name
    elif (hour >=5  and hour < 12): signoff = 'That is all for this morning. Have a pleasant day'       + signoff_name
    elif (hour >=12 and hour < 17): signoff = 'That is all for this afternoon. Have a pleasant evening' + signoff_name
    else:                           signoff = 'That is all for this evening. Have a pleasant night'     + signoff_name
    return signoff

def get_date():
    short_time = time.strftime("%I:%M%p", time.localtime())
    long_date = time.strftime("%A, %B %d", time.localtime())
    return "It's " + short_time + ' on ' + long_date + ". "

def get_weather():
    weather = pywapi.get_weather_from_yahoo(location_id, 'metric')
    current_condition = string.lower(weather['condition']['text'])
    current_temp = weather['condition']['temp']
    high_temp = weather['forecasts'][0]['high']
    low_temp = weather['forecasts'][0]['low']
    return 'The weather in ' + city + ' is currently ' + current_condition + ' with a temperature of ' + current_temp + ' degrees. '\
           'Today there will be a high of ' + high_temp + ', and a low of ' + low_temp + '. '

def get_newmail_count():
    connection = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    connection.login(email_username, email_password)
    unreadCount = int(re.search("UNSEEN (\d+)", connection.status("INBOX", "(UNSEEN)")[1][0]).group(1))
    if   (unreadCount == 0):    result = 'You do not have any new e-mails. '
    elif (unreadCount == 1):    result = 'You have one new e-mail. '
    else:                       result = 'You have ' + str(unreadCount) + ' new e-mails. '    
    return result

def main():
    voice = win32com.client.Dispatch("SAPI.SpVoice")
    text = get_greeting() + get_date() + get_weather() + get_newmail_count() + get_signoff()
    print text
    voice.Speak(text)
    
if __name__ == '__main__':
    main()