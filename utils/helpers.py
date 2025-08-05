import random
import string
from config import error_messages
from helium import Text
from config import stop_flag


def generate_password():
    lowercase = ''.join(random.choices(string.ascii_lowercase, k=4))
    uppercase = ''.join(random.choices(string.ascii_uppercase, k=3))
    digits = ''.join(random.choices(string.digits, k=7))
    password = list(lowercase + uppercase + digits)
    random.shuffle(password)
    return ''.join(password)

def get_gender_title(title):
    return "Mr" if title == "Mr." else "Ms"




def check_for_errors(row, password, results, user_id):
    try:
        for error_msg, error_text in error_messages.items():
            if Text(error_text).exists():
                print(f"Ошибка на странице: {error_text}")
                results.append(row + [password, user_id, "BAD", error_msg])
                return True  # Возвращаем True, если ошибка найдена
        return False  # Возвращаем False, если ошибок нет
    except Exception as e:
        print(f"Ошибка в функции check_for_errors: {e}")
        return False  # Возвращаем False, чтобы не прерывать выполнение
