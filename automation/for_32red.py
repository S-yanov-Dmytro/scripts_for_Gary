import time
import keyboard
import threading
import pyautogui
import subprocess

from helium import *
from config import sites, month_map
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


def run_automation_32red(data_list, file_path, lbl_status):
    results = []
    listener_thread = threading.Thread(target=listen_for_exit_key, daemon=True)
    listener_thread.start()

    for row in data_list:
        error_recorded = False
        password = generate_password()
        PAGE_URL = sites["32red"]
        driver = None
        reconnect_vpn()

        try:
            driver = start_chrome(PAGE_URL)
            title, first_name, last_name, _, address, town, _, city, county, postcode, mobile, email, dob_date = row
            title = get_gender_title_betvictor(title)
            year, month, day = dob_date.split('-')
            time.sleep(4)
            click("Allow all cookies")
            time.sleep(1)
            click('Join the fun')
            time.sleep(2)
            write(first_name, into="First name")
            write(last_name, into='Last name')
            write(Keys.PAGE_DOWN)
            time.sleep(2)
            print(month)
            select("Day", str(day))
            select("Month", month_map[month])
            select('Year', str(year))
            time.sleep(2)
            write(email, into='Email')
            write(password, into='Password')
            select("Select a gender", title)
            time.sleep(1)
            if check_for_errors(row, password, results, user_id='...'):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Current iteration completed, script will be stopped.")
                    return
                raise Exception("Error after first page")
            click('Continue')
            time.sleep(2)
            address_clk = f"{address} {postcode}"

            write(address_clk, into="Address")
            wait_until(lambda: find_all(S('[data-test-name^="lookup-result-"]')))
            result = find_all(S('[data-test-name^="lookup-result-"]'))
            click(result[0])
            time.sleep(1)
            write(Keys.PAGE_DOWN)

            write(mobile, into='Mobile Number')
            time.sleep(1)
            click("Yes")
            time.sleep(1)
            click("Casino")
            write(Keys.ARROW_DOWN)
            time.sleep(1)
            click("SMS")
            write(Keys.ARROW_DOWN)
            time.sleep(1)
            click("I am 18 or over and I accept the Terms and Conditions.")
            write(Keys.ARROW_DOWN)
            time.sleep(3)

            click("Join")
            time.sleep(8)
            if Text("HI, DO WE KNOW YOU?").exists():
                results.append(row + [password, '...', "BAD", "Duplicate"])
                print("Duplicate")
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Current iteration completed, script will be stopped.")
                    return
            write(Keys.PAGE_DOWN)
            click("Opt in to your bonus")
            time.sleep(2)



            if Text("Verify your identity & address").exists():
                results.append(row + [password, '...', "CNV", "Verify Identity"])
                print("Verify Identity")
            elif Text("You’re in!").exists():
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