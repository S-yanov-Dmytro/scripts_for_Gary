import time
import keyboard
import threading
import pyautogui
import logging
from datetime import datetime

from helium import *
from config import sites
from utils.helpers import generate_password, get_gender_title_pari
from utils.file_operations import save_results_to_excel
from utils.browser_utils import kill_chrome
from utils.helpers import check_for_errors

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('automation.log')
    ]
)
logger = logging.getLogger(__name__)

stop_flag = False


def listen_for_exit_key():
    global stop_flag
    keyboard.add_hotkey('ctrl+q', lambda: set_stop_flag())
    while not stop_flag:
        time.sleep(0.1)


def set_stop_flag():
    global stop_flag
    stop_flag = True
    logger.info("⛔ Зупинка скрипта ініційована користувачем (Ctrl+Q)")


def run_automation_parimatch(data_list, file_path, lbl_status):
    results = []
    listener_thread = threading.Thread(target=listen_for_exit_key, daemon=True)
    listener_thread.start()

    for idx, row in enumerate(data_list, 1):
        if stop_flag:
            logger.info("Отримано сигнал зупинки. Завершення роботи.")
            save_results_to_excel(results, file_path)
            return

        error_recorded = False
        password = generate_password()
        PAGE_URL = sites["parimatch"]
        driver = None

        try:
            logger.info(f"🚀 Початок обробки запису {idx}/{len(data_list)}")
            driver = start_chrome(PAGE_URL)
            logger.info("Браузер успішно запущено")

            title, first_name, last_name, _, address, town, _, city, county, postcode, mobile, email, dob_date = row
            title = get_gender_title_pari(title)
            year, month, day = dob_date.split('-')

            # Логируем основные данные
            logger.debug(f"Дані: {first_name} {last_name}, {dob_date}, {mobile}, {email}")

            # Первая страница
            logger.info("Заповнення першої сторінки реєстрації")
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
                    logger.info("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                    return
                raise Exception("Ошибка после первой страницы")

            click('Continue')
            logger.info("Перша сторінка успішно заповнена")

            # Вторая страница
            logger.info("Заповнення другої сторінки реєстрації")
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
                    logger.info("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                    return
                raise Exception("Ошибка после второй страницы")

            click("Continue To Last Step")
            logger.info("Друга сторінка успішно заповнена")
            time.sleep(2)

            # Третья страница
            logger.info("Заповнення третьої сторінки реєстрації")
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
            logger.info("Форма успішно відправлена, очікування відповіді...")
            time.sleep(3)

            if check_for_errors(row, password, results, user_id='...'):
                if stop_flag:
                    save_results_to_excel(results, file_path)
                    logger.info("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                    return
                raise Exception("Ошибка после третьей страницы")

            click("Submit")
            logger.info("Submit натиснуто, очікування 10 секунд...")

            start_wait = datetime.now()
            time.sleep(10)
            logger.info(f"Час очікування минув ({(datetime.now() - start_wait).total_seconds()} сек)")

            try:
                logger.info("Спроба отримати статус реєстрації...")
                title_element = S("h3.bvs-msg-box__title")
                title_text = title_element.web_element.get_attribute("innerText").strip()
                logger.info(f"🟢 Заголовок модального окна: {repr(title_text)}")
            except Exception as e:
                title_text = ""
                logger.error(f"⛔ Не удалось получить заголовок: {e}")

            if Text("Account Created").exists():
                results.append(row + [password, "OK", "Success"])
                logger.info("✅ Успішна реєстрація: Account Created")
            elif Text("Duplicate Account").exists():
                results.append(row + [password, "BAD", "Duplicate"])
                logger.warning("⚠ Дублікат акаунта: Duplicate Account")
            elif Text("Success").exists():
                results.append(row + [password, "OK", "Success"])
                logger.info("✅ Успішна реєстрація: Success")
            elif Text("Player Verification").exists():
                results.append(row + [password, "CNV", "Player Verification"])
                logger.warning("⚠ Потрібна верифікація: Player Verification")
            elif title_text == "Verify Your Account":
                results.append(row + [password, "CNV", "Verify Identity"])
                logger.warning("⚠ Потрібна верифікація: Verify Your Account")
            else:
                results.append(row + [password, "BAD", "Reg failed"])
                logger.error("❌ Помилка реєстрації: невідома відповідь")

            if stop_flag:
                logger.info("✅ Поточна ітерація завершена, скрипт буде зупинено.")
                save_results_to_excel(results, file_path)
                return

        except Exception as e:
            logger.error(f"❌ Виникла помилка: {str(e)}", exc_info=True)
            if not error_recorded:
                save_results_to_excel(results, file_path)
                error_recorded = True
        finally:
            if driver:
                try:
                    driver.quit()
                    logger.info("Браузер успішно закрито")
                except Exception as e:
                    logger.error(f"Помилка при закритті браузера: {e}")
            kill_chrome()

    save_results_to_excel(results, file_path)
    lbl_status.config(text="All done!", fg="green")
    logger.info("✅ Всі операції успішно завершені")