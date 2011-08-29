import win32com.client
import pywapi
import random
import string
import time

first_name = 'Adam'
last_name = 'Mansfield'
city = 'hamilton'
location_id = 'CAXX0184'

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
    currentCondition = string.lower(weather['condition']['text'])
    currentTemp = weather['condition']['temp']
    highTemp = weather['forecasts'][0]['high']
    lowTemp = weather['forecasts'][0]['low']
    return 'The weather in ' + city + ' is currently ' + currentCondition + ' with a temperature of ' + currentTemp + ' degrees. '\
           'Today there will be a high of ' + highTemp + ', and a low of ' + lowTemp + '. '

def main():
    voice = win32com.client.Dispatch("SAPI.SpVoice")
    text = get_greeting() + get_date() + get_weather() + get_signoff()
    print text
    voice.Speak(text)

if __name__ == '__main__':
    main()