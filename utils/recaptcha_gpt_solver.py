from captcha_solver import solve_captcha_with_gpt
from selenium.webdriver.firefox.options import Options
from selenium import webdriver
import os
from dotenv import load_dotenv

load_dotenv()

def handle_puzzle_captcha():
    try:
        print("[GPT-4] Handling image reCAPTCHA via GPT-4...")
        solve_captcha_with_gpt()
        print("[GPT-4] Puzzle CAPTCHA solved.")
    except Exception as e:
        print(f"[GPT-4] Error solving puzzle captcha: {e}")
