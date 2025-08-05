import re
import time
import keyboard
import threading
import pyautogui
import subprocess

from helium import *
from config import sites

from utils.browser_utils import kill_chrome
from utils.file_operations import save_results_to_excel
from utils.helpers import generate_password, get_gender_title, check_for_errors

stop_flag = False

def listen_for_exit_key():
    global stop_flag
    keyboard.add_hotkey('ctrl+q', lambda: set_stop_flag())
    keyboard.wait('ctrl+q')


def set_stop_flag():
    global stop_flag
    stop_flag = True
    print("\n⛔ Зупинка скрипта ініційована користувачем (Ctrl+Q)\n")



def open_expressvpn():
    print("⏳start ExpressVPN...")
    subprocess.Popen([r"C:\Program Files (x86)\ExpressVPN\expressvpn-ui\ExpressVPN.exe"])
    time.sleep(3)


def reconnect_vpn():
    open_expressvpn()
    print("⏳ OFF VPN...")
    pyautogui.click(930, 540)
    time.sleep(3)
    print("⏳ ON VPN...")
    pyautogui.click(930, 540)
    time.sleep(10)
    print("✅ VPN reconnected!")


def run_automation_betway(data_list, file_path, lbl_status=None):
    global stop_flag
    results = []
    listener_thread = threading.Thread(target=listen_for_exit_key, daemon=True)
    listener_thread.start()

    for row in data_list:
        error_recorded = False
        driver = None
        reconnect_vpn()
        password = generate_password()
        PAGE_URL = sites["betway"]
        try:
            driver = start_chrome(PAGE_URL)

            time.sleep(4)

            title, first_name, last_name, _, address, _, _, city, county, postcode, mobile, email, dob_date = row
            title = get_gender_title(title)
            year, month, day = dob_date.split('-')
            user_id = first_name[:3].lower() + str(mobile)[-4:] + "q12e"

            click("Reject all")
            time.sleep(1)

            click("Join")
            print(stop_flag)

            time.sleep(1)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)

            if Text("casino").exists():
                click("casino")
                time.sleep(1)
                if check_for_errors(row, password, results, user_id):
                    if stop_flag:
                        save_results_to_excel(results, file_path)
                        print("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                        return
                    raise Exception("Ошибка после выбора Vegas")
                click("Next")
            else:
                click("Change offer")
                time.sleep(1)
                click("Stake")
                click("Continue")

            
            time.sleep(1)
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                    return
                raise Exception("Ошибка после первого Next")

            select('Title', title)
            write(first_name, into='First Name')
            write(last_name, into='Last Name')
            select("dd", str(int(day)))
            select("mm", str(month))
            select('yyyy', str(year))
            time.sleep(1)
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                    return
                raise Exception("Ошибка после ввода личных данных")

            click('Next')
            time.sleep(1)
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                    return
                raise Exception("Ошибка после второго Next")

            write(user_id, into='Username')
            write(password, into='Password')
            write(email, into='Email')
            time.sleep(1)
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                    return
                raise Exception("Ошибка после логина/почты")

            click('Next')
            time.sleep(1)
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                    return
                raise Exception("Ошибка после третьего Next")

            click("or enter address manually")
            time.sleep(1)
            write(address, into='Street Address')
            write(city, into='City')
            write(postcode, into='Post Code')
            time.sleep(1)
            select("Please select", "Wales")
            write(mobile, into='Mobile')
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                    return
                raise Exception("Ошибка после адреса и телефона")

            select("Daily limit", '£100')
            select("Weekly limit", '£100')
            select("Monthly limit", '£100')
            click(S("#Comp17_TermsAndConditions"))
            click("Email")
            time.sleep(2)
            click('Register')
            time.sleep(2)
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                    return
                raise Exception("Ошибка после регистрации")

            time.sleep(8)

            try:
                el = S("span.sc-dbdfbe3d-0")
                text_el = el.web_element.get_attribute("innerText")
                text_el = normalize_text(text_el)
                print(f"✅ Текст из span.sc-dbdfbe3d-0: {repr(text_el)}")
            except Exception as e:
                text_el = ""
                print(f"⛔ Не удалось получить text_el: {e}")

            # Получаем текст из второго span
            try:
                element = S("span.ng-binding")
                text = element.web_element.get_attribute("innerText")
                text = normalize_text(text)
                print(f"👉 Найдено в span.ng-binding: {repr(text)}")
            except Exception as e:
                text = ""
                print(f"⛔ Не удалось получить text: {e}")

            # Получаем текст модального окна
            try:
                modal_body = S(".modal-body").web_element.text
                modal_body = normalize_text(modal_body)
                print("📦 Содержимое модального окна:\n", repr(modal_body))
            except Exception as e:
                modal_body = ""
                print(f"⛔ Не удалось получить modal_body: {e}")

            # Проверка условий
            if "Before you can deposit or stake, we need to verify your identity." in text_el:
                results.append(row + [password, user_id, "CNV", "Verification Failed"])
                print("Verification Failed")
            elif "Choose a deposit method" in text:
                results.append(row + [password, user_id, "OK", "Success"])
                print("Success")
            elif "you already have an account" in modal_body:
                results.append(row + [password, user_id, "BAD", "Duplicate"])
                print("Duplicate")
            elif "verify your identity" in text_el:
                results.append(row + [password, user_id, "CNV", "Verification Failed"])
                print("Verification Failed")
            else:
                print('⚠️ Не удалось определить результат — неизвестный случай')


            if stop_flag:
                print("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                save_results_to_excel(results, file_path)
                return

        except Exception as e:
            print(e)
            if not error_recorded:
                save_results_to_excel(results, file_path)
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
            kill_chrome()

    save_results_to_excel(results, file_path)
    lbl_status.config(text="All done!", fg="green")

def normalize_text(text):
    """Удаляет лишние пробелы, переносы, табы и т.п."""
    return re.sub(r'\s+', ' ', text).strip()