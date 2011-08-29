import win32com.client
import time
import random


first_name = 'Adam'
last_name = 'Mansfield'


def get_greeting():
    option = random.uniform(0,100)
    if   (option<=25): greetingName = ', Mr. ' + last_name + '. '
    elif (option<=50): greetingName = ' ' + first_name + '. '
    elif (option<=75): greetingName = ' sir. '
    else:              greetingName = '. '

    hour = int(time.strftime("%H", time.localtime()))
    option = random.uniform(0,100)
    if   (option<=20):              greeting = 'Hello'          + greetingName
    elif (option<=40):              greeting = 'Hi'             + greetingName
    elif (hour >=0 and hour < 5):   greeting = 'Good night'     + greetingName
    elif (hour >=5 and hour < 12):  greeting = 'Good morning'   + greetingName
    elif (hour >=12 and hour < 17): greeting = 'Good afternoon' + greetingName
    else:                           greeting = 'Good evening'   + greetingName
    return greeting

def get_signoff():
    option = random.uniform(0,100)
    if   (option<=33): signoffName = ' ' + first_name + '. '
    elif (option<=66): signoffName = ' sir. '
    else:              signoffName = '. '

    hour = int(time.strftime("%H", time.localtime()))
    if   (hour >=0  and hour < 5):  signoff = 'That is all for tonight. Have a pleasant sleep'          + signoffName
    elif (hour >=5  and hour < 12): signoff = 'That is all for this morning. Have a pleasant day'       + signoffName
    elif (hour >=12 and hour < 17): signoff = 'That is all for this afternoon. Have a pleasant evening' + signoffName
    else:                           signoff = 'That is all for this evening. Have a pleasant night'     + signoffName
    return signoff

def get_date():
    short_time = time.strftime("%I:%M%p", time.localtime())
    long_date = time.strftime("%A, %B %d", time.localtime())
    return "It's " + short_time + ' on ' + long_date + ". "

def main():
    voice = win32com.client.Dispatch("SAPI.SpVoice")
    text = get_greeting() + get_date() + get_signoff()
    print text
    voice.Speak(text)

if __name__ == '__main__':
    main()