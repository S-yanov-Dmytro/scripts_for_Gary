from selenium import webdriver
from helium import *
import time
from config import sites, error_messages
from utils.helpers import generate_password, get_gender_title
from utils.file_operations import save_results_to_excel
from utils.browser_utils import kill_chrome


def check_for_errors(row, password, results):
    """Проверка наличия текстовых ошибок на странице"""
    for err_key, err_text in error_messages.items():
        if Text(err_text).exists():
            results.append(row + [password, "BAD", err_key])
            raise Exception(err_key)


def run_automation_ladbrokes(data_list, file_path, lbl_status):
    results = []

    for row in data_list:
        driver = None
        error_recorded = False
        password = generate_password()

        try:
            options = webdriver.ChromeOptions()
            driver = webdriver.Chrome(options=options)
            set_driver(driver)
            start_chrome(sites["Ladbrokes"])  # Заменить на нужный сайт при необходимости

            title, first_name, last_name, _, address, _, _, city, county, postcode, mobile, email, dob_date = row
            title = get_gender_title(title)
            time.sleep(3)

            click('JOIN HERE')
            time.sleep(2)
            check_for_errors(row, password, results)

            write(email, into='Email')
            time.sleep(2)
            check_for_errors(row, password, results)

            user_id = first_name[:3].lower() + str(mobile)[-4:]
            write(user_id, into='User ID')
            time.sleep(1)
            check_for_errors(row, password, results)

            write(password, into='Create password')
            time.sleep(2)
            check_for_errors(row, password, results)

            if Text("Necessary Only").exists():
                click('Necessary Only')
                time.sleep(1)

            click('CONTINUE')
            time.sleep(2)
            check_for_errors(row, password, results)

            click(title)
            write(first_name, into='First name')
            write(last_name, into='Last name')
            time.sleep(1)
            check_for_errors(row, password, results)

            year, month, day = dob_date.split('-')
            write(day, into='DD')
            write(month, into='MM')
            write(year, into='YYYY')
            time.sleep(1)

            click('CONTINUE')
            time.sleep(2)
            check_for_errors(row, password, results)

            write(mobile, into='Mobile number')
            time.sleep(1)
            click('Enter address manually')
            time.sleep(1)

            write(address, into='House No. and Street Name')
            write(city, into='City')
            write(postcode, into='Postcode')
            time.sleep(1)
            check_for_errors(row, password, results)

            click('Select all options')
            time.sleep(1)
            click('CREATE MY ACCOUNT')
            time.sleep(7)
            check_for_errors(row, password, results)

            if Text("Verification Failed").exists():
                results.append(row + [password, "CNV", "Verification Failed"])
            else:
                results.append(row + [password, "OK", "Success"])

        except Exception as e:
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
