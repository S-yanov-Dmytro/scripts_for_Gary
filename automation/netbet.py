import requests
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementClickInterceptedException

import time
import os
import math
import tempfile
import traceback
import re
import keyboard
import threading
import subprocess
import pyautogui
import pandas as pd



from helium import *
from utils.helpers import check_for_errors, generate_password, get_gender_title
from utils.file_operations import save_results_to_excel
from utils.browser_utils import kill_chrome
from config import sites
from openai import OpenAI
from PIL import Image
from dotenv import load_dotenv
from selenium.common.exceptions import TimeoutException, NoSuchElementException

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
    print("⏳ Відкриваємо ExpressVPN...")
    subprocess.Popen([r"C:\Program Files (x86)\ExpressVPN\expressvpn-ui\ExpressVPN.exe"])
    time.sleep(3)  # Чекаємо, поки програма відкриється


def reconnect_vpn():
    open_expressvpn()
    print("⏳ Відключаємо VPN...")
    pyautogui.click(930, 540)
    time.sleep(5)
    print("⏳ Підключаємо VPN знову...")
    pyautogui.click(930, 540)
    time.sleep(5)
    print("✅ VPN перепідключено!")

load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

IMGUR_CLIENT_ID = os.getenv("IMGUR_CLIENT_ID", "ae8db954c4defd4")



def upload_image_to_imgur(image_path):
    print(f"[Imgur] Uploading image {image_path}...")
    url = "https://api.imgur.com/3/image"
    headers = {"Authorization": f"Client-ID {IMGUR_CLIENT_ID}"}
    with open(image_path, "rb") as f:
        response = requests.post(url, headers=headers, files={"image": f})
    res = response.json()
    if res.get("success"):
        return res["data"]["link"]
    else:
        raise Exception(f"[Imgur] Upload failed: {res}")

def ask_recaptcha_whole_view(image_path, challenge_text, tile_count, client):
    grid_size = int(math.sqrt(tile_count))
    if grid_size * grid_size != tile_count:
        return []

    try:
        full_img_url = upload_image_to_imgur(image_path)
    except Exception as e:
        print("[Imgur] Failed to upload full image:", e)
        return []

    prompt = (
        f"You are shown an image divided into a {grid_size}x{grid_size} grid of tiles, numbered from 0 to {tile_count - 1}.\n"
        f"The task is to select only the tiles that clearly contain the following object:\n\n"
        f"\"{challenge_text.strip()}\"\n\n"
        f"Instructions:\n"
        f"- Select tiles ONLY if they visibly contain the specified object.\n"
        f"- Do NOT select tiles that are empty, blurry, or contain unrelated objects.\n"
        f"- Provide your answer as a comma-separated list of tile numbers (e.g., 0,3,5).\n"
        f"- If no tiles match, reply with \"none\".\n"
    )
    messages = [
        {"role": "system", "content": "Reply only with tile numbers. No explanation."},
        {"role": "user", "content": [
            {"type": "text", "text": prompt},
            {"type": "image_url", "image_url": {"url": full_img_url, "detail": "high"}}
        ]}
    ]

    try:
        resp = client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            max_tokens=256,
            temperature=0,
        )
        answer = resp.choices[0].message.content.strip().lower()
        tiles = list(map(int, re.findall(r'\b\d+\b', answer)))
        return tiles
    except Exception as e:
        print("[OpenAI] GPT error:", e)
        return []

def is_recaptcha_images_challenge_present(driver, timeout=5):
    try:
        iframe = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "iframe[title='recaptcha challenge expires in two minutes']"))
        )
        return True
    except (TimeoutException, NoSuchElementException):
        return False


def solve_recaptcha_images_only(driver, max_attempts=15):
    print("[CAPTCHA] Starting OpenAI reCAPTCHA solver...")
    for attempt in range(1, max_attempts + 1):
        print(f"[CAPTCHA] Attempt #{attempt}")
        try:
            iframe = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[src*='bframe']"))
            )
            driver.switch_to.frame(iframe)

            instructions = WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.rc-imageselect-instructions"))
            )

            tile_wrappers = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.rc-image-tile-wrapper"))
            )
            panel = driver.find_element(By.CSS_SELECTOR, "div.rc-imageselect-target")

            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_file:
                captcha_img_path = tmp_file.name
                panel.screenshot(captcha_img_path)

            challenge_text = instructions[0].text

            tiles_to_click = ask_recaptcha_whole_view(captcha_img_path, challenge_text, len(tile_wrappers), client)
            os.remove(captcha_img_path)

            for tile_idx in tiles_to_click:
                if 0 <= tile_idx < len(tile_wrappers):
                    ActionChains(driver).move_to_element(tile_wrappers[tile_idx]).click().perform()
                    time.sleep(0.8)

            verify_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "recaptcha-verify-button"))
            )
            verify_button.click()
            time.sleep(3)

            driver.switch_to.default_content()
            time.sleep(2)

            # Перевірка чи капча зникла
            if not is_recaptcha_images_challenge_present(driver):
                print('[CAPTCHA] Капча успішно пройдена!')
                return True

        except Exception:
            print(f"[CAPTCHA] Помилка в спробі #{attempt}:\n{traceback.format_exc()}")
            driver.switch_to.default_content()
            time.sleep(2)

    print("[CAPTCHA] Всі спроби вичерпані, капча не пройдена.")
    return False




def wait_until_clear(driver, row, password, results):
    while True:
        if is_recaptcha_images_challenge_present(driver):
            if not solve_recaptcha_images_only(driver):
                return False
        if check_for_errors(row, password, results):
            return False
        if not is_recaptcha_images_challenge_present(driver):
            return True


def run_automation_netbet(data_list, file_path, lbl_status):
    listener_thread = threading.Thread(target=listen_for_exit_key, daemon=True)
    listener_thread.start()
    results = []
    PAGE_URL = sites["netbet"]

    temp_path = file_path.replace(".xlsx", "_temp_results.xlsx")
    if os.path.exists(temp_path):
        try:
            df = pd.read_excel(temp_path, header=None)
            results = df.where(pd.notnull(df), "").values.tolist()
            print(f"⚠️ Обнаружен временный файл с результатами. Загружено {len(results)} записей.")
        except Exception as e:
            print(f"⛔ Ошибка загрузки временного файла: {e}")

    try:
        for idx, row in enumerate(data_list, start=1):
            if stop_flag:
                print("⛔ Обнаружен запрос на остановку скрипта")
                save_results_to_excel(results, file_path)
                lbl_status.config(text="Stopped by user", fg="red")
                return

            if any(r[:len(row)] == list(row) for r in results):
                print(f"[{idx}/{len(data_list)}] Пропускаем уже обработанную запись")
                continue

            reconnect_vpn()
            driver = None
            password = generate_password()

            try:
                print(f"[{idx}/{len(data_list)}] Opening page...")
                driver = start_chrome(PAGE_URL)

                title, first_name, last_name, _, address, _, _, city, county, postcode, mobile, email, dob_date = row
                title = get_gender_title(title)
                user_id = first_name[:3].lower() + str(mobile)[-4:]

                click("Register")
                time.sleep(2)
                if Text("What country").exists():
                    click("Continue")

                write(email, into='E-MAIL')
                write(mobile, into='NUMBER')
                write(user_id, into='USER')
                write(Keys.PAGE_DOWN)
                time.sleep(1)
                click("Continue")
                time.sleep(2)
                if is_recaptcha_images_challenge_present(driver):
                    if not solve_recaptcha_images_only(driver):
                        print("Не вдалося вирішити капчу. Завершую сесію.")
                        results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                        driver.quit()
                        kill_chrome()
                        continue
                else:
                    print("Капча відсутня, продовжуємо.")

                if check_for_errors(row, password, results):
                    continue


                time.sleep(1)

                write(password, into='PASSWORD')
                time.sleep(1)

                while True:
                    # 1. Капча
                    if is_recaptcha_images_challenge_present(driver):
                        if not solve_recaptcha_images_only(driver):
                            print("Не вдалося вирішити капчу. Завершую сесію.")
                            results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                            driver.quit()
                            kill_chrome()
                            continue
                        time.sleep(1)

                    # 2. Помилки
                    if check_for_errors(row, password, results):
                        print("Помилка на сторінці після капчі.")
                        results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                        driver.quit()
                        kill_chrome()
                        continue

                    # 3. Все чисто
                    if not is_recaptcha_images_challenge_present(driver) and not check_for_errors(row, password, results):
                        break  # Усе ок — йдемо далі

                click("Continue")

                write(first_name, into='FIRST NAME')
                time.sleep(1)
                write(last_name, into='LAST NAME')
                write(Keys.PAGE_DOWN)

                while True:
                    # 1. Капча
                    if is_recaptcha_images_challenge_present(driver):
                        if not solve_recaptcha_images_only(driver):
                            print("Не вдалося вирішити капчу. Завершую сесію.")
                            results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                            driver.quit()
                            kill_chrome()
                            continue
                        time.sleep(1)

                    # 2. Помилки
                    if check_for_errors(row, password, results):
                        print("Помилка на сторінці після капчі.")
                        results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                        driver.quit()
                        kill_chrome()
                        continue

                    # 3. Все чисто
                    if not is_recaptcha_images_challenge_present(driver) and not check_for_errors(row, password, results):
                        break  # Усе ок — йдемо далі
                time.sleep(1)

                click("TITLE")
                write(Keys.PAGE_DOWN)
                time.sleep(1)
                click(title)

                while True:
                    # 1. Капча
                    if is_recaptcha_images_challenge_present(driver):
                        if not solve_recaptcha_images_only(driver):
                            print("Не вдалося вирішити капчу. Завершую сесію.")
                            results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                            driver.quit()
                            kill_chrome()
                            continue
                        time.sleep(1)

                    # 2. Помилки
                    if check_for_errors(row, password, results):
                        print("Помилка на сторінці після капчі.")
                        results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                        driver.quit()
                        kill_chrome()
                        continue

                    # 3. Все чисто
                    if not is_recaptcha_images_challenge_present(driver) and not check_for_errors(row, password, results):
                        break  # Усе ок — йдемо далі
                day, month, year = dob_date.split('-')
                date = f"{year}{month}{day}"
                write(date, into='DATE OF BIRTH')

                while True:
                    # 1. Капча
                    if is_recaptcha_images_challenge_present(driver):
                        if not solve_recaptcha_images_only(driver):
                            print("Не вдалося вирішити капчу. Завершую сесію.")
                            results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                            driver.quit()
                            kill_chrome()
                            continue
                        time.sleep(1)

                    # 2. Помилки
                    if check_for_errors(row, password, results):
                        print("Помилка на сторінці після капчі.")
                        results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                        driver.quit()
                        kill_chrome()
                        continue

                    # 3. Все чисто
                    if not is_recaptcha_images_challenge_present(driver) and not check_for_errors(row, password, results):
                        break  # Усе ок — йдемо далі

                write(Keys.PAGE_DOWN)
                time.sleep(1)
                click("Continue")

                correct_address = f"{address}, {city}, {postcode}"
                print(f"[Address] Filling in: {correct_address}")
                write(correct_address, into='TYPE IN YOUR ADDRESS')
                time.sleep(3)
                write(Keys.ENTER)

                while True:
                    # 1. Капча
                    if is_recaptcha_images_challenge_present(driver):
                        if not solve_recaptcha_images_only(driver):
                            print("Не вдалося вирішити капчу. Завершую сесію.")
                            results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                            driver.quit()
                            kill_chrome()
                            continue
                        time.sleep(1)

                    # 2. Помилки
                    if check_for_errors(row, password, results):
                        print("Помилка на сторінці після капчі.")
                        results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                        driver.quit()
                        kill_chrome()
                        continue

                    # 3. Все чисто
                    if not is_recaptcha_images_challenge_present(driver) and not check_for_errors(row, password, results):
                        break  # Усе ок — йдемо далі
                write(Keys.PAGE_DOWN)
                time.sleep(1)
                write(Keys.PAGE_DOWN)
                time.sleep(1)
                click("Continue")
                time.sleep(2)

                click(S("#option2 + label"))
                time.sleep(1)
                write(Keys.PAGE_DOWN)
                time.sleep(1)
                click(S("#check1 + label"))
                time.sleep(1)
                click("Start playing")

                while True:
                    # 1. Капча
                    if is_recaptcha_images_challenge_present(driver):
                        if not solve_recaptcha_images_only(driver):
                            print("Не вдалося вирішити капчу. Завершую сесію.")
                            results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                            driver.quit()
                            kill_chrome()
                            continue
                        time.sleep(1)

                    # 2. Помилки
                    if check_for_errors(row, password, results):
                        print("Помилка на сторінці після капчі.")
                        results.append(row + [password, user_id, "BAD", "Reg failed-Сaptcha", first_name, last_name])
                        driver.quit()
                        kill_chrome()
                        continue

                    # 3. Все чисто
                    if not is_recaptcha_images_challenge_present(driver) and not check_for_errors(row, password, results):
                        break  # Усе ок — йдемо далі

                time.sleep(10)
                click("NO")
                time.sleep(5)

                if Text("Proof of ID").exists():
                    results.append(row + [password, "CNV", "Verification Failed"])
                else:
                    results.append(row + [password, "OK", "Success"])

            except Exception as e:
                print(e)
                results.append(row + [password, "BAD", str(e)])

            finally:
                if driver:
                    try:
                        driver.quit()
                    except Exception as e:
                        print(f"[{idx}] Error quitting driver: {e}")
                kill_chrome()
    except Exception as e:
        print(f"⛔ Критическая ошибка: {str(e)}")
        traceback.print_exc()
    finally:
        save_results_to_excel(results, file_path)
        lbl_status.config(text="All done!", fg="green")
