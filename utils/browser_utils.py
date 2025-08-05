import subprocess

def kill_chrome():
    try:
        subprocess.call("taskkill /F /IM chrome.exe", shell=True,
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.call("taskkill /F /IM chromedriver.exe", shell=True,
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        print(f"Error while closing Chrome: {e}")
