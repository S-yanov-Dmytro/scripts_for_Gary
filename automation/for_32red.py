import time
import random
import threading
import keyboard
import pyautogui
import undetected_chromedriver as uc
from helium import set_driver, click, write, Text, select, Keys, wait_until, find_all, kill_browser, S
from config import sites, month_map
from utils.helpers import generate_password, get_gender_title_betvictor, check_for_errors
from utils.file_operations import save_results_to_excel
from utils.browser_utils import kill_chrome

stop_flag = False


def listen_for_exit_key():
    keyboard.add_hotkey('ctrl+q', lambda: set_stop_flag())
    keyboard.wait('ctrl+q')


def set_stop_flag():
    global stop_flag
    stop_flag = True
    print("\n⛔ Script stop initiated by user (Ctrl+Q)\n")


def random_sleep(a=1, b=3):
    time.sleep(random.uniform(a, b))


def start_chrome_with_proxy(url):
    options = uc.ChromeOptions()
    proxy_host = "proxy.smartproxy.net"
    proxy_port = 3120
    options.add_argument(f"--proxy-server={proxy_host}:{proxy_port}")

    # запускаем без fullscreen
    driver = uc.Chrome(options=options)
    set_driver(driver)
    driver.get(url)

    # Устанавливаем нужный размер окна и позицию
    driver.set_window_rect(x=0, y=0, width=1100, height=1050)

    return driver


def run_automation_32red(data_list, file_path, lbl_status):
    results = []
    listener_thread = threading.Thread(target=listen_for_exit_key, daemon=True)
    listener_thread.start()

    for row in data_list:
        password = generate_password()
        PAGE_URL = sites["32red"]
        driver = None

        try:
            driver = start_chrome_with_proxy(PAGE_URL)
            random_sleep(7, 9)

            # Пример имитации ввода
            pyautogui.write("smart-bnlx5pj03qrj_area-GB")
            random_sleep(1, 2)
            pyautogui.click(650, 350)
            random_sleep(1, 2)
            pyautogui.write("5Hzwnwx2sZ4BvYbt")
            random_sleep(1, 2)
            pyautogui.click(735, 435)
            random_sleep(3, 5)

            title, first_name, last_name, _, address, town, _, city, county, postcode, mobile, email, dob_date = row
            title = get_gender_title_betvictor(title)
            year, month, day = dob_date.split('-')

            wait_until(lambda: Text("Allow all cookies").exists())

            click("Allow all cookies")
            random_sleep(1, 2)

            click('Register')
            random_sleep(2, 3)
            write(first_name, into="First name")
            write(last_name, into="Last name")
            write(Keys.PAGE_DOWN)
            random_sleep(1, 2)

            select("Day", str(day))
            select("Month", month_map[month])
            select('Year', str(year))
            random_sleep(1, 2)

            write(email, into='Email')
            write(password, into='Password')
            select("Select a gender", title)
            random_sleep(1, 2)

            click('Continue')
            random_sleep(2, 3)

            address_input = f"{address} {postcode}"
            write(address_input, into="Address")
            wait_until(lambda: find_all(S('[data-test-name^="lookup-result-"]')))
            results_found = find_all(S('[data-test-name^="lookup-result-"]'))
            click(results_found[0])
            random_sleep(1, 2)

            write(mobile, into='Mobile Number')
            random_sleep(1, 2)
            click("Yes")
            random_sleep(1, 2)
            click("Casino")
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            random_sleep(1, 2)
            click("SMS")
            for _ in range(7):
                write(Keys.ARROW_DOWN)
            random_sleep(1, 2)
            click("I am 18 or over and I accept the Terms and Conditions.")
            random_sleep(2, 3)
            click("Join")
            random_sleep(5, 7)

            if Text("Your account is blocked").exists():
                results.append(row + [password, '...', "CNV", "Verify Identity"])
                print("Verify Identity")
            elif Text("HI, DO WE KNOW YOU?").exists():
                results.append(row + [password, '...', "BAD", "Duplicate"])
                print("Duplicate")
            else:
                click("Opt in to your bonus")
                random_sleep(2, 3)
                if Text("You’re in!").exists():
                    results.append(row + [password, '...', "OK", "Success"])
                    print("Success")
                else:
                    results.append(row + [password, '...', "BAD", "Reg failed"])
                    print("Reg failed")

            if stop_flag:
                save_results_to_excel(results, file_path)
                return

        except Exception as e:
            print(e)
            save_results_to_excel(results, file_path)
        finally:
            if driver:
                try:
                    kill_browser()
                except Exception:
                    pass
            kill_chrome()

    save_results_to_excel(results, file_path)
    lbl_status.config(text="All done!", fg="green")
