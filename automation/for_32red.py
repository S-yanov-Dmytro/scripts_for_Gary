import time
import keyboard
import threading
import pyautogui
from helium import start_chrome, click, write, Text, select, Keys, find_all, S, wait_until, kill_browser
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
    set_driver(driver)  # Устанавливаем созданный драйвер для Helium
    driver.get(url)
    return driver


def run_automation_32red(data_list, file_path, lbl_status):
    results = []
    listener_thread = threading.Thread(target=listen_for_exit_key, daemon=True)
    listener_thread.start()

    for row in data_list:
        error_recorded = False
        password = generate_password()
        PAGE_URL = sites["32red"]
        driver = None

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

            title, first_name, last_name, _, address, town, _, city, county, postcode, mobile, email, dob_date = row
            title = get_gender_title_betvictor(title)
            year, month, day = dob_date.split('-')

            if Text("Allow all cookies").exists():
                click("Allow all cookies")
            time.sleep(1)

            click('Register')
            time.sleep(2)

            write(first_name, into="First name")
            write(last_name, into='Last name')
            write(Keys.PAGE_DOWN)
            time.sleep(1)

            select("Day", str(day))
            select("Month", month_map[month])
            select('Year', str(year))
            time.sleep(1)

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

            address_input = f"{address} {postcode}"
            write(address_input, into="Address")
            wait_until(lambda: find_all(S('[data-test-name^="lookup-result-"]')))
            result = find_all(S('[data-test-name^="lookup-result-"]'))
            click(result[0])
            time.sleep(1)

            write(mobile, into='Mobile Number')
            time.sleep(1)
            click("Yes")
            time.sleep(1)
            click("Casino")
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)

            time.sleep(1)
            click("SMS")
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)
            write(Keys.ARROW_DOWN)

            time.sleep(1)
            click("I am 18 or over and I accept the Terms and Conditions.")
            time.sleep(2)
            click("Join")
            time.sleep(7)

            if Text("HI, DO WE KNOW YOU?").exists():
                results.append(row + [password, '...', "BAD", "Duplicate"])
                print("Duplicate")
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    return

            elif Text("Your account is blocked").exists():
                results.append(row + [password, '...', "CNV", "Verify Identity"])
                print("Verify Identity")
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    return
            else:
                write(Keys.PAGE_DOWN)
                click("Opt in to your bonus")
                time.sleep(2)

                if Text("You’re in!").exists():
                    results.append(row + [password, '...', "OK", "Success"])
                    print("Success")
                else:
                    results.append(row + [password, '...', "BAD", "Reg failed"])
                    print("Reg failed")

                if stop_flag:
                    save_results_to_excel(results, file_path)
                    print("✅ Script stopped by user.")
                    return

        except Exception as e:
            print(e)
            if not error_recorded:
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
