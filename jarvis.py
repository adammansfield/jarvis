import win32com.client
import time


def get_date():
    return "It's " + time.strftime("%I:%M%p", time.localtime()) + ' on ' + time.strftime("%A, %B %d", time.localtime()) + ". "


def main():
    voice = win32com.client.Dispatch("SAPI.SpVoice")
    text = get_date()
    print text
    voice.Speak(text)


if __name__ == '__main__':
    main()