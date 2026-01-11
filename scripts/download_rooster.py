import os
import time
import logging
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# --- CONFIGURATION ---
# I have cleaned the Proofpoint wrapper from your URL and added 'download=1' to force a file download
TARGET_URL = "https://amspecllc-my.sharepoint.com/:x:/r/personal/okan_ozturk_amspecgroup_com/Documents/Desktop/Rooster%202026.xlsm?download=1"
OUTPUT_FILENAME = "Rooster 2026.xlsm"
AUTH_FILE = "auth.json"  # Stores your login cookies
LOG_FILE = "logs/download_log.txt"

# Ensure logs directory exists
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
    """Helper to ensure consistent user feedback"""
    if type == "SUCCESS":
        logging.info(f"✅ SUCCESS: {message}")
        print(f"\033[92m✅ SUCCESS: {message}\033[0m") # Green text
    elif type == "ERROR":
        logging.error(f"❌ ERROR: {message}")
        print(f"\033[91m❌ ERROR: {message}\033[0m") # Red text
    elif type == "INFO":
        logging.info(f"ℹ️  INFO: {message}")
        print(f"ℹ️  INFO: {message}")

def run():
    with sync_playwright() as p:
        log_feedback("INFO", "Initializing Browser...")

        # We launch headless=False so you can see the login screen if needed
        # On M1 Macs and Windows, Chromium is generally the most stable for this
        browser = p.chromium.launch(headless=False)

        # Load auth state if it exists
        context_args = {}
        if os.path.exists(AUTH_FILE):
            log_feedback("INFO", "Found existing authentication state.")
            context_args["storage_state"] = AUTH_FILE
        else:
            log_feedback("INFO", "No authentication state found. You will need to log in manually.")

        context = browser.new_context(**context_args)
        page = context.new_page()

        try:
            log_feedback("INFO", f"Navigating to SharePoint...")

            # Start the download by navigating to the URL
            # We use expect_download to grab the event
            try:
                with page.expect_download(timeout=60000) as download_info:
                    page.goto(TARGET_URL)

                    # DETECT LOGIN SCREEN:
                    # If we are redirected to a login page (microsoftonline.com), wait for user
                    if "login.microsoftonline.com" in page.url:
                        log_feedback("INFO", "Login required! Please log in inside the browser window.")
                        log_feedback("INFO", "Script is pausing for 120 seconds to allow you to complete MFA/Login...")

                        # Wait for the user to finish login and for the URL to redirect back to SharePoint
                        # We wait for the URL to contain the original sharepoint domain
                        page.wait_for_url("https://amspecllc-my.sharepoint.com/**", timeout=120000)

                        # Save the new state so next time works automatically
                        context.storage_state(path=AUTH_FILE)
                        log_feedback("SUCCESS", "Authentication saved to auth.json")

                download = download_info.value

                # Overwrite the local file
                download.save_as(OUTPUT_FILENAME)
                log_feedback("SUCCESS", f"File downloaded and saved as: {os.path.abspath(OUTPUT_FILENAME)}")

            except PlaywrightTimeoutError:
                log_feedback("ERROR", "Download timed out. Did you log in successfully?")
                raise # Re-raise to signal failure to caller
            except Exception as e:
                log_feedback("ERROR", f"An unexpected error occurred during download: {e}")
                raise

        except Exception as e:
            log_feedback("ERROR", f"Browser automation failed: {e}")
            raise
        finally:
            browser.close()

if __name__ == "__main__":
    run()
