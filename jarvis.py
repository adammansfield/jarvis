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


# userconfig variables 
first_name = 'John'
last_name = 'Doe'
city = 'New York'
location_id = '10001'
email_username = 'default@gmail.com'
email_password = '1234'

# import custom userconfig if it exists
try:
    from userconfig import *
except ImportError:
    if (os.path.exists(os.getcwd() + "\\config\\default.py")):
        shutil.copyfile(os.getcwd() + "\\config\\default.py", os.getcwd() + "\\config\\userconfig.py")


def get_greeting():
    """Return a randomized, personalized greeting."""
    name_chance = random.uniform(0,100)
    greeting_chance = random.uniform(0,100)
    
    if (name_chance <= 25): 
        name = ', Mr. ' + last_name + '. '
    elif (name_chance <= 50):
        name = ' ' + first_name + '. '
    elif (name_chance <= 75):
        name = ' sir. '
    else:
        name = '. '

    if (greeting_chance <= 15):
        greeting = 'Hello' + name
    elif (greeting_chance <= 30):
        greeting = 'Hi' + name
    else:
        hour = int(time.strftime("%H", time.localtime()))
        if (hour >=0  and hour < 5):
            greeting = 'Good night' + name
        elif (hour >=5  and hour < 12):
            greeting = 'Good morning' + name
        elif (hour >=12 and hour < 17):
            greeting = 'Good afternoon' + name
        else:
            greeting = 'Good evening' + name
            
    return greeting

def get_signoff():
    """Return a sign off depending on the time of day."""
    name_chance = random.uniform(0,100)
    
    if (name_chance <= 33):
        name = ' ' + first_name + '. '
    elif (name_chance <= 66):
        name = ' sir. '
    else:
        name = '. '

    hour = int(time.strftime("%H", time.localtime()))
    if (hour >= 0  and hour < 5):
        signoff = 'That is all for tonight. Have a pleasant sleep' + name
    elif (hour >= 5  and hour < 12):
        signoff = 'That is all for this morning. Have a pleasant day' + name
    elif (hour >= 12 and hour < 17):
        signoff = 'That is all for this afternoon. Have a pleasant evening' + name
    else:
        signoff = 'That is all for this evening. Have a pleasant night' + name
        
    return signoff

def get_date():
    """Return the time (12 hour) and date (weekday, month day)."""
    short_time = time.strftime("%I:%M%p", time.localtime())
    long_date = time.strftime("%A, %B %d", time.localtime())
    
    return "It's " + short_time + ' on ' + long_date + ". "

def get_weather():
    """Return the weather temperature and conditions for the day."""
    weather = pywapi.get_weather_from_yahoo(location_id, 'metric')
    
    current_condition = string.lower(weather['condition']['text'])
    current_temp = weather['condition']['temp']
    high_temp = weather['forecasts'][0]['high']
    low_temp = weather['forecasts'][0]['low']
    
    return 'The weather in ' + city + ' is currently ' + current_condition + \
           ' with a temperature of ' + current_temp + ' degrees. Today there ' + \
           'will be a high of ' + high_temp + ', and a low of ' + low_temp + '. '

def get_newmail_count():
    """Return the number of new email from gmail."""
    # if default email then return an example output
    if (email_username == "default@gmail.com"):
        return 'You have 7 new e-mails. '
    
    connection = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    connection.login(email_username, email_password)
    
    unreadCount = int(re.search("UNSEEN (\d+)", 
                      connection.status("INBOX", "(UNSEEN)")[1][0]).group(1))
    
    if (unreadCount == 0):
        result = 'You do not have any new e-mails. '
    elif (unreadCount == 1):
        result = 'You have one new e-mail. '
    else:
        result = 'You have ' + str(unreadCount) + ' new e-mails. '    
        
    return result


def main():
    """Concats and speaks all the information."""
    voice = win32com.client.Dispatch("SAPI.SpVoice")
    speech = get_greeting() + get_date() + get_weather() + get_newmail_count() + get_signoff()
    print speech
    voice.Speak(speech)
    
if __name__ == '__main__':
    main()