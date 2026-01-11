import os
import time
import logging
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# --- CONFIGURATION ---
BASE_URL = "https://amspecllc-my.sharepoint.com/:x:/r/personal/okan_ozturk_amspecgroup_com/Documents/Desktop/Rooster%202026.xlsm?d=wf6d964fbb614486688c587499e010aaa&e=4%3a4c933ba0b5f2418abbb97d622c2865ec&sharingv2=true&fromShare=true"

OUTPUT_FILENAME = "Rooster  2026.xlsm"
AUTH_FILE = "auth.json"
LOG_FILE = "logs/download_log.txt"

os.makedirs("logs", exist_ok=True)

# --- LOGGING SETUP ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)

def log_feedback(type, message):
    if type == "SUCCESS":
        logging.info(f"SUCCESS: {message}")
        print(f"SUCCESS: {message}")
    elif type == "ERROR":
        logging.error(f"ERROR: {message}")
        print(f"ERROR: {message}")
    elif type == "INFO":
        logging.info(f"INFO: {message}")
        print(f"INFO: {message}")

def run():
    with sync_playwright() as p:
        log_feedback("INFO", "Initializing Browser...")
        browser = p.chromium.launch(headless=False)

        context_args = {}
        if os.path.exists(AUTH_FILE):
            log_feedback("INFO", "Found existing authentication state.")
            context_args["storage_state"] = AUTH_FILE

        context = browser.new_context(**context_args)
        page = context.new_page()

        try:
            # --- STEP 1: LOAD VIEWER ---
            log_feedback("INFO", "Step 1: Opening Excel Viewer...")
            page.goto(BASE_URL)

            # Auth Check
            if "login.microsoftonline.com" in page.url:
                log_feedback("INFO", "Login required. Please log in.")
                try:
                    page.wait_for_url("https://amspecllc-my.sharepoint.com/**", timeout=120000)
                    log_feedback("SUCCESS", "Login detected. Saving state...")
                    context.storage_state(path=AUTH_FILE)
                except PlaywrightTimeoutError:
                    return

            log_feedback("INFO", "Waiting for session to stabilize...")
            page.wait_for_load_state("domcontentloaded")
            time.sleep(10) # Wait for JS to initialize

            # --- STEP 2: EXTRACT DYNAMIC DOWNLOAD LINK ---
            log_feedback("INFO", "Step 2: Extracting dynamic download token...")

            # Microsoft stores the download URL with a temporary auth token in a global variable
            # We execute JS to retrieve it.
            download_url = page.evaluate("""() => {
                if (typeof _wopiContextJson !== 'undefined' && _wopiContextJson.FileGetUrl) {
                    return _wopiContextJson.FileGetUrl;
                }
                return null;
            }""")

            if not download_url:
                log_feedback("ERROR", "Could not find dynamic download URL in page variables.")
                # Fallback: Try to find it in the HTML text if JS variable is hidden/renamed
                content = page.content()
                import re
                match = re.search(r'"FileGetUrl":"(https:[^"]+)"', content)
                if match:
                    download_url = match.group(1).replace(r'\u0026', '&')
                    log_feedback("INFO", "Found URL via Regex fallback.")
                else:
                    raise Exception("Failed to extract download URL.")

            log_feedback("INFO", "Found dynamic URL. Triggering download...")

            # --- STEP 3: DOWNLOAD ---
            with page.expect_download(timeout=60000) as download_info:
                # We use window.open to respect the auth context
                page.evaluate(f"window.open('{download_url}')")

            download = download_info.value
            download.save_as(OUTPUT_FILENAME)

            # Verify file size
            size = os.path.getsize(OUTPUT_FILENAME)
            if size < 2000: # If smaller than 2KB, it's likely an HTML error page
                 raise Exception(f"Downloaded file is too small ({size} bytes). Likely an error page.")

            log_feedback("SUCCESS", f"File downloaded: {OUTPUT_FILENAME} ({size} bytes)")

        except Exception as e:
            log_feedback("ERROR", f"Process failed: {e}")
            # Debug snapshot
            page.screenshot(path="logs/error_final.png")
            raise
        finally:
            browser.close()

if __name__ == "__main__":
    run()
