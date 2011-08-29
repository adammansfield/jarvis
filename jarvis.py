import time


def get_date():
    return "It's " + time.strftime("%I:%M%p", time.localtime()) + ' on ' + time.strftime("%A, %B %d", time.localtime()) + ". "


def main():
    text = get_date()
    print text


if __name__ == '__main__':
    main()