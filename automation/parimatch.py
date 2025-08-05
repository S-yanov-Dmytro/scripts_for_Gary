import time
import pyautogui

from helium import *
from config import sites, error_messages
from utils.helpers import generate_password, get_gender_title
from utils.file_operations import save_results_to_excel
from utils.browser_utils import kill_chrome

from utils.helpers import check_for_errors


def run_automation_parimatch(data_list, file_path, lbl_status):
    results = []

    for row in data_list:
        error_recorded = False
        password = generate_password()
        PAGE_URL = sites["parimatch"]
        driver = None

        try:
            driver = start_chrome(PAGE_URL)

            title, first_name, last_name, _, address, town, _, city, county, postcode, mobile, email, dob_date = row
            title = get_gender_title(title)
            year, month, day = dob_date.split('-')
            time.sleep(3)
            pyautogui.moveTo(822, 960, duration=0.2)
            pyautogui.click()
            time.sleep(1)

            click('Sign Up')
            time.sleep(2)
            # if check_for_errors(row, password, results):
            #     continue
            click(title)
            write(first_name, into='First name')
            write(last_name, into='Last name')
            time.sleep(2)
            write(f'{day}{month}{year}', into="Date of Birth")
            time.sleep(2)
            write(mobile, into='Mobile Number')
            time.sleep(2)
            click('Continue')

            # if check_for_errors(row, password, results):
            #     continue
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
            time.sleep(1)
            click(S('button.regv2-button-submit'))
            time.sleep(10)

            if Text("Verification Failed").exists():
                results.append(row + [password, "CNV", "Verification Failed"])
            else:
                results.append(row + [password, "OK", "Success"])

        except Exception as e:
            print(e)
            if not error_recorded:
                results.append(row + [password, "BAD", str(e)])
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
            kill_chrome()

    save_results_to_excel(results, file_path)
    lbl_status.config(text="All done!", fg="green")
