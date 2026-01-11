import subprocess
import sys
import os
import logging
from datetime import datetime

# --- CONFIGURATION ---
PYTHON_CMD = sys.executable  # Use current python interpreter
SCRIPT_DIR = "scripts"
DOWNLOAD_SCRIPT = os.path.join(SCRIPT_DIR, "download_rooster.py")
SYNC_SCRIPT = os.path.join(SCRIPT_DIR, "rooster_sync.py")
LOG_FILE = "logs/workflow.log"

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

def run_command(command, description):
    """Run a shell command and return success status."""
    logging.info(f"Starting: {description}")
    try:
        result = subprocess.run(command, check=True, text=True, capture_output=True)
        logging.info(f"Completed: {description}")
        if result.stdout:
            logging.debug(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed: {description}")
        logging.error(f"Error Output: {e.stderr}")
        return False

def git_operations():
    """Handle git add, commit, push."""
    logging.info("Checking for git changes...")

    # Check status
    status = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True)
    if not status.stdout.strip():
        logging.info("No changes to commit.")
        return True

    logging.info("Changes detected. Committing and pushing...")

    if not run_command(["git", "add", "output/calendars/*.ics"], "Git Add Calendars"):
        return False

    commit_msg = f"Auto-update schedules: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    if not run_command(["git", "commit", "-m", commit_msg], "Git Commit"):
        return False

    if not run_command(["git", "push"], "Git Push"):
        return False

    return True

def main():
    logging.info("="*50)
    logging.info("Starting Daily Rooster Workflow")
    logging.info("="*50)

    # 1. Download Excel
    if not run_command([PYTHON_CMD, DOWNLOAD_SCRIPT], "Download Excel File"):
        logging.error("Workflow aborted due to download failure.")
        sys.exit(1)

    # 2. Sync Calendars
    if not run_command([PYTHON_CMD, SYNC_SCRIPT], "Generate Calendars"):
        logging.error("Workflow aborted due to sync failure.")
        sys.exit(1)

    # 3. Git Push
    if not git_operations():
        logging.error("Workflow failed during git operations.")
        sys.exit(1)

    logging.info("Workflow completed successfully.")

if __name__ == "__main__":
    main()
