import re
import time
import keyboard
import threading
import pyautogui

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
    print("\n⛔ Script stop initiated by user (Ctrl+Q)\n")


def start_chrome_with_proxy(url):
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from webdriver_manager.chrome import ChromeDriverManager
    from helium import set_driver  # Импортируем set_driver из Helium

    proxy_host = "proxy.smartproxy.net"
    proxy_port = 3120

    options = Options()
    options.add_argument(f"--proxy-server={proxy_host}:{proxy_port}")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    set_driver(driver)
    driver.get(url)
    return driver


def run_automation_betway(data_list, file_path, lbl_status=None):
    global stop_flag
    results = []
    listener_thread = threading.Thread(target=listen_for_exit_key, daemon=True)
    listener_thread.start()

    for row in data_list:
        error_recorded = False
        driver = None
        password = generate_password()
        PAGE_URL = sites["betway"]
        try:
            driver = start_chrome_with_proxy(PAGE_URL)
            time.sleep(3)
            pyautogui.write("smart-bnlx5pj03qrj_area-GB_state-england")
            time.sleep(1)
            pyautogui.click(650, 350)
            time.sleep(1)
            pyautogui.write("5Hzwnwx2sZ4BvYbt")
            time.sleep(1)
            pyautogui.click(735, 435)
            time.sleep(13)

            title, first_name, last_name, _, address, _, _, city, county, postcode, mobile, email, dob_date = row
            title = get_gender_title(title)
            year, month, day = dob_date.split('-')
            user_id = first_name[:3].lower() + str(mobile)[-4:] + "q12e00"

            wait_until(lambda: Text("Reject all").exists())
            click("Reject all")
            time.sleep(1)

            click("Join")

            time.sleep(3)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)

            if Text("casino").exists():
                click("casino")
                time.sleep(1)
                if check_for_errors(row, password, results, user_id):
                    if stop_flag:
                        save_results_to_excel(results, file_path)
                        print("✅ Current iteration completed, script will be stopped.")
                        return
                    raise Exception("Error after selecting Vegas")
                click("Next")
            else:
                click("Change offer")
                time.sleep(1)
                click("Stake")
                click("Continue")

            time.sleep(2)
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Current iteration completed, script will be stopped.")
                    return
                raise Exception("Error after first Next")

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
                    print("✅ Current iteration completed, script will be stopped.")
                    return
                raise Exception("Error after entering personal data")
            if Text("Reject all").exists():
                click("Reject all")

            click('Next')
            time.sleep(1)
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Current iteration completed, script will be stopped.")
                    return
                raise Exception("Error after second Next")

            write(user_id, into='Username')
            write(password, into='Password')
            write(email, into='Email')
            time.sleep(1)
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Current iteration completed, script will be stopped.")
                    return
                raise Exception("Error after login/email")

            click('Next')
            time.sleep(1)
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Current iteration completed, script will be stopped.")
                    return
                raise Exception("Error after third Next")

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
                    print("✅ Current iteration completed, script will be stopped.")
                    return
                raise Exception("Error after address and phone")

            select("Daily limit", '£100')
            select("Weekly limit", '£100')
            select("Monthly limit", '£100')
            click(S("#Comp17_TermsAndConditions"))
            click("Email")
            time.sleep(2)
            click('Register')
            time.sleep(1)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            time.sleep(5)
            if check_for_errors(row, password, results, user_id):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Current iteration completed, script will be stopped.")
                    return
                raise Exception("Error after registration")

            time.sleep(8)

            try:
                el = S("span.sc-dbdfbe3d-0")
                text_el = el.web_element.get_attribute("innerText")
                text_el = normalize_text(text_el)
                print(f"✅ Text from span.sc-dbdfbe3d-0: {repr(text_el)}")
            except Exception as e:
                text_el = ""
                print(f"⛔ Failed to get text_el: {e}")

            try:
                element = S("span.ng-binding")
                text = element.web_element.get_attribute("innerText")
                text = normalize_text(text)
                print(f"👉 Found in span.ng-binding: {repr(text)}")
            except Exception as e:
                text = ""
                print(f"⛔ Failed to get text: {e}")

            try:
                modal_body = S(".modal-body").web_element.text
                modal_body = normalize_text(modal_body)
                print("📦 Modal window content:\n", repr(modal_body))
            except Exception as e:
                modal_body = ""
                print(f"⛔ Failed to get modal_body: {e}")

            # Check conditions
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
                print('⚠️ Failed to determine result - unknown case')

            if stop_flag:
                print("✅ Current iteration completed, script will be stopped.")
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
    return re.sub(r'\s+', ' ', text).strip()