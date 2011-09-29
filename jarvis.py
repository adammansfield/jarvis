import win32com.client
import pywapi
import imaplib
import os
import random
import re
import shutil
import string
import sys
import time
import urllib2
sys.path.append(os.getcwd() + "\\config")


# default config variables
config_title = 'Mr.' 
config_first_name = 'John'
config_last_name = 'Doe'
config_formal_address = 'sir'
config_email_username = 'default@gmail.com'
config_email_password = 'hunter2'

# try to import custom config variables
try:
    from userconfig import *
except ImportError:
    if os.path.exists(os.getcwd() + '\\config\\default.py'):
        shutil.copyfile(os.getcwd() + '\\config\\default.py', os.getcwd() + '\\config\\userconfig.py')


def get_name(first_name, last_name, title, formal_address):
    """Return a radomized name using title, address, and first or last name"""
    chance = random.uniform(0, 100)
    
    if chance < 25: 
        name = title + ' ' + last_name
    elif chance < 50:
        name = first_name
    elif chance < 75:
        name = formal_address
    else:
        name = ''
        
    return name

def get_greeting():
    """Return a randomized, personalized greeting."""
    chance = random.uniform(0, 100)
    
    if chance <= 15:
        greeting = 'Hello'
    elif chance <= 30:
        greeting = 'Hi'
    else:
        hour = int(time.strftime('%H', time.localtime()))
        if hour >=0  and hour < 5:
            greeting = 'Good night'
        elif hour >=5  and hour < 12:
            greeting = 'Good morning'
        elif hour >=12 and hour < 17:
            greeting = 'Good afternoon'
        else:
            greeting = 'Good evening'
    
    return greeting

def get_signoff():
    """Return a sign off depending on the time of day."""
    hour = int(time.strftime('%H', time.localtime()))
    if hour >= 0  and hour < 5:
        signoff = 'That is all for tonight. Have a pleasant sleep'
    elif hour >= 5  and hour < 12:
        signoff = 'That is all for this morning. Have a pleasant day'
    elif hour >= 12 and hour < 17:
        signoff = 'That is all for this afternoon. Have a pleasant evening'
    else:
        signoff = 'That is all for this evening. Have a pleasant night'
    
    return signoff

def get_date():
    """Return the time (12 hour) and date (weekday, month day)."""
    short_time = time.strftime('%I:%M%p', time.localtime())
    long_date = time.strftime('%A, %B %d', time.localtime())
    
    return "It's " + short_time + ' on ' + long_date + '. '

def get_weather():
    """Return the weather temperature and conditions for the day."""
    geodata = urllib2.urlopen("http://j.maxmind.com/app/geoip.js").read()
    country_result = re.split("country_name\(\) \{ return \'", geodata)
    country_result = re.split("'; }", country_result[1])
    city_result = re.split("city\(\)         \{ return \'", geodata)
    city_result = re.split("'; }", city_result[1])
    country = country_result[0]
    city = city_result[0]
    
    weather = pywapi.get_weather_from_google(city, country)
    
    if not weather['current_conditions']:
        print 'ERROR: Get weather failed.\n'
        return ''
    
    current_condition = string.lower(weather['current_conditions']['condition'])
    current_temp = weather['current_conditions']['temp_c']
    high_temp_f = weather['forecasts'][0]['high']
    low_temp_f = weather['forecasts'][0]['low']
    
    high_temp_c = int((int(high_temp_f) - 32) / 1.8)
    low_temp_c = int((int(low_temp_f) - 32) / 1.8)
    
    return 'The weather in ' + city + ' is currently ' + current_condition + \
           ' with a temperature of ' + current_temp + ' degrees. Today there ' + \
           'will be a high of ' + str(high_temp_c) + ', and a low of ' + str(low_temp_c) + '. '

def get_newmail_count():
    """Return the number of new email from gmail."""
    try:
        connection = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        connection.login(email_username, email_password)
    except:
        print 'ERROR: Gmail login failed.\n'
        return ''
    
    unreadCount = int(re.search("UNSEEN (\d+)", 
                      connection.status("INBOX", "(UNSEEN)")[1][0]).group(1))
    
    if unreadCount is 0:
        result = 'You do not have any new e-mails. '
    elif unreadCount is 1:
        result = 'You have one new e-mail. '
    else:
        result = 'You have ' + str(unreadCount) + ' new e-mails. '    
        
    return result


def main():
    """Concat and speak all the information."""
    voice = win32com.client.Dispatch("SAPI.SpVoice")
    speech = \
        get_greeting() + ' ' + \
        get_name(config_title, config_first_name, config_last_name, config_formal_address) + '. ' + \
        get_date() + \
        get_weather() + \
        get_newmail_count() + \
        get_signoff() + ' ' + \
        get_name(config_title, config_first_name, config_last_name, config_formal_address) + '.' 
    print speech
    voice.Speak(speech)
    
if __name__ == '__main__':
    main()
