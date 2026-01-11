# Rooster Sync - Windows Lab PC Setup Guide

This guide describes how to set up the Rooster Sync automation on a fresh Windows PC (Lab Computer).

## Prerequisites
*   Administrator access to the PC.
*   Internet connection.
*   Chocolatey installed (as per your description).

## 1. Install Software (via PowerShell Administrator)

Open PowerShell as Administrator and run:

```powershell
# Install Git
choco install git -y

# Install Python (ensure it adds to PATH)
choco install python -y
```

*Close and reopen PowerShell after installation to refresh the PATH.*

## 2. Clone Repository & Setup Project

Navigate to where you want the project to live (e.g., Documents):

```powershell
cd $HOME\Documents
git clone https://github.com/kelvinalexander/Rooster.git
cd Rooster
```

## 3. Python Environment Setup

Create the virtual environment and install dependencies:

```powershell
# Create virtual environment
python -m venv venv

# Activate it
.\venv\Scripts\Activate

# Install dependencies (including Playwright)
pip install -r requirements.txt

# Install Playwright browsers
playwright install chromium
```

## 4. GitHub Authentication (First Time Only)

To allow the script to push changes automatically, you need to authenticate Git.

1.  **Configure User:**
    ```powershell
    git config --global user.name "Kelvin Alexander"
    git config --global user.email "your.email@example.com"
    ```
2.  **Authenticate:**
    We recommend using the Git Credential Manager (installed with Git).
    When you run the first `git push`, a window will pop up asking you to sign in to GitHub.
    *   Sign in with your browser.
    *   This saves a token so you don't need to log in again.

## 5. SharePoint Authentication (First Time Run)

You need to run the download script manually once to handle the Microsoft 2FA login.

1.  Make sure your virtual environment is active (`.\venv\Scripts\Activate`).
2.  Run the download script:
    ```powershell
    python scripts/download_rooster.py
    ```
3.  **Action:** The Chrome browser will open and navigate to the Excel file.
4.  **Action:** Log in to your Microsoft account. Approve the MFA on your phone.
5.  **Wait:** Once logged in, the script will automatically detect the session, download the file, and **close the browser itself**. Do not close it manually unless it gets stuck for more than 2 minutes.

*Verification:* Check if `auth.json` exists in the project folder.

## 6. Automate with Task Scheduler

We will set up the script to run automatically at 08:50 and 16:00.

1.  Open **Task Scheduler**.
2.  Click **Create Basic Task**.
3.  **Name:** `RoosterSync Morning`.
4.  **Trigger:** Daily > Start time: `08:50:00`.
5.  **Action:** Start a program.
    *   **Program/script:** `path\to\your\venv\Scripts\python.exe` (Find absolute path via `Get-Command python`)
    *   **Add arguments:** `run_workflow.py`
    *   **Start in:** `C:\Users\YourUser\Documents\Rooster` (The project folder path)
6.  Finish.
7.  Repeat for `RoosterSync Afternoon` at `16:00:00`.

**Important:** In the Task Properties, verify:
*   "Run only when user is logged on" (Required for Headed browser execution, though we might switch to headless if `auth.json` works reliably).
*   If you want it to run while locked, you might need to change the Playwright script to `headless=True` in `scripts/download_rooster.py` after the initial setup.

## 7. Troubleshooting

*   **Logs:** Check `logs/workflow.log` for success/failure messages.
*   **Browser showing up:** If the browser pops up every day, edit `scripts/download_rooster.py` and change `headless=False` to `headless=True`.
