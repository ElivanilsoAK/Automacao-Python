import os
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Load Config
def load_config():
    try:
        with open("config.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

config = load_config()

# Setup Driver (Visible mode for debugging)
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--start-maximized")
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-software-rasterizer")
chrome_options.add_argument("window-size=1920,1080")
chrome_options.add_argument("--log-level=3") # Minimal logging from chrome


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

try:
    print("Navigating to InControl...")
    driver.get(config.get("incontrol_url", ""))
    time.sleep(5)

    # Login
    print("Attempting Login...")
    try:
        if driver.find_elements(By.CSS_SELECTOR, "input[type='password']"):
            user = config.get("incontrol_user", "")
            password = config.get("incontrol_password", "")
            
            driver.find_element(By.CSS_SELECTOR, "input[type='text'], input[name='username']").send_keys(user)
            driver.find_element(By.CSS_SELECTOR, "input[type='password']").send_keys(password)
            driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
            print("Login submitted. Waiting...")
            time.sleep(10)
    except Exception as e:
        print(f"Login step error: {e}")

    # Ensure we are on the report page
    if "eventos-usuario" not in driver.current_url:
        print("Redirecting to events page...")
        driver.get(config.get("incontrol_url", ""))
        time.sleep(8)

    print("Capturing Page Source and Screenshot...")
    
    # Save HTML
    with open("debug_page_source.html", "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    
    # Save Screenshot
    driver.save_screenshot("debug_screenshot.png")
    
    print("DONE! Files 'debug_page_source.html' and 'debug_screenshot.png' saved.")
    print("Please inspect these files or share them.")

except Exception as e:
    print(f"Critical Error: {e}")
finally:
    driver.quit()
