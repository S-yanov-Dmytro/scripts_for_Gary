import random
import string
import keyboard
from config import error_messages
from helium import Text

stop_flag = False

def listen_for_exit_key():
    global stop_flag
    keyboard.add_hotkey('ctrl+q', lambda: set_stop_flag())
    keyboard.wait('ctrl+q')

def set_stop_flag():
    global stop_flag
    stop_flag = True
    print("\n⛔ Зупинка скрипта ініційована користувачем (Ctrl+Q)\n")


def generate_password():
    lowercase = ''.join(random.choices(string.ascii_lowercase, k=4))
    uppercase = ''.join(random.choices(string.ascii_uppercase, k=3))
    digits = ''.join(random.choices(string.digits, k=7))
    password = list(lowercase + uppercase + digits)
    random.shuffle(password)
    return ''.join(password)

def get_gender_title(title):
    return "Mr" if title == "Mr." else "Ms"

def get_gender_title_betvictor(title):
    return "Male" if title == "Mr." else "Female"


def check_for_errors(row, password, results, user_id):
    try:
        for error_msg, error_text in error_messages.items():
            if Text(error_text).exists():
                print(error_text)
                print(f"Error on page: {error_text}")
                results.append(row + [password, user_id, "BAD", error_msg])
                return True
        return False
    except Exception as e:
        print(f"Error check_for_errors: {e}")
        return False

