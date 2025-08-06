import time
import keyboard
import threading
import pyautogui
import subprocess

from helium import *
from config import sites
from utils.helpers import generate_password, get_gender_title_betvictor
from utils.file_operations import save_results_to_excel
from utils.browser_utils import kill_chrome

from utils.helpers import check_for_errors


stop_flag = False

def listen_for_exit_key():
    global stop_flag
    keyboard.add_hotkey('ctrl+q', lambda: set_stop_flag())
    keyboard.wait('ctrl+q')

def set_stop_flag():
    global stop_flag
    stop_flag = True
    print("\n⛔ Script stop initiated by user (Ctrl+Q)\n")

def open_expressvpn():
    print("⏳ Starting ExpressVPN...")
    subprocess.Popen([r"C:\Program Files (x86)\ExpressVPN\expressvpn-ui\ExpressVPN.exe"])
    time.sleep(3)


def reconnect_vpn():
    open_expressvpn()
    print("⏳ Turning OFF VPN...")
    pyautogui.click(930, 540)
    time.sleep(3)
    print("⏳ Turning ON VPN...")
    pyautogui.click(930, 540)
    time.sleep(10)
    print("✅ VPN reconnected!")


def run_automation_betvictor(data_list, file_path, lbl_status):
    results = []
    listener_thread = threading.Thread(target=listen_for_exit_key, daemon=True)
    listener_thread.start()

    for row in data_list:
        error_recorded = False
        password = generate_password()
        PAGE_URL = sites["betvictor"]
        driver = None
        reconnect_vpn()

        try:
            driver = start_chrome(PAGE_URL)
            title, first_name, last_name, _, address, town, _, city, county, postcode, mobile, email, dob_date = row
            title = get_gender_title_betvictor(title)
            year, month, day = dob_date.split('-')
            time.sleep(3)
            pyautogui.moveTo(822, 960, duration=0.2)
            pyautogui.click()

            time.sleep(1)
            click('Sign Up')
            time.sleep(2)
            click(title)
            write(first_name, into='First name')
            write(last_name, into='Last name')
            time.sleep(2)
            write(f'{day}{month}{year}', into="Date of Birth")
            time.sleep(2)
            write(mobile, into='Mobile Number')
            time.sleep(2)
            if check_for_errors(row, password, results, user_id='...'):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Current iteration completed, script will be stopped.")
                    return
                raise Exception("Error after first page")
            click('Continue')

            time.sleep(1)
            write(address, into='Address Line 1')
            time.sleep(1)
            write(town, into='Town / Village')
            time.sleep(1)
            write(city, into="City")
            time.sleep(1)
            write(county, into="County")
            time.sleep(1)
            write(postcode, into='Postcode')
            time.sleep(1)
            if check_for_errors(row, password, results, user_id='...'):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Current iteration completed, script will be stopped.")
                    return
                raise Exception("Error after second page")
            click("Continue To Last Step")
            time.sleep(2)

            write(email, into='Email')
            write(password, into='Password')
            write(Keys.PAGE_DOWN)
            time.sleep(1)
            click(S('label[for="CASINO-EMAIL"]'))

            click(S('label[for="SPORT-EMAIL"]'))
            time.sleep(1)
            write(Keys.PAGE_DOWN)
            click(S('label[for="termsPrivacyPolicy"]'))
            time.sleep(2)
            click(S('button.regv2-button-submit'))

            time.sleep(5)
            if check_for_errors(row, password, results, user_id='...'):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Current iteration completed, script will be stopped.")
                    return
                raise Exception("Error after third page")

            try:
                title_element = S("h3.bvs-msg-box__title")
                title_text = title_element.web_element.get_attribute("innerText").strip()
                print(f"🟢 Modal window title: {repr(title_text)}")
            except Exception as e:
                title_text = ""
                print(f"⛔ Failed to get title: {e}")
            if title_text == "Verify Your Account":
                results.append(row + [password, '...', "CNV", "Verify Identity"])
                print("Verify Identity")
            elif Text("Account Created").exists():
                results.append(row + [password, '...', "OK", "Success"])
                print("Success1")
            else:
                results.append(row + [password, '...', "BAD", "Reg failed"])
                print("Reg failed")

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