# Rooster Sync Deployment Checklist

**Goal:** Get the automated calendar sync running on the Windows Lab PC.

## Phase 1: Environment & Code
- [ ] **Open Terminal:** Launch PowerShell as Administrator.
- [ ] **Install Prerequisites:**
    ```powershell
    choco install git python -y
    ```
    *(Close and reopen PowerShell after this)*
- [ ] **Clone Repository:**
    ```powershell
    cd $HOME\Documents
    git clone https://github.com/kelalexandermurphy/Rooster.git
    cd Rooster
    ```
- [ ] **Setup Python:**
    ```powershell
    python -m venv venv
    .\venv\Scripts\Activate
    pip install -r requirements.txt
    playwright install chromium
    ```

## Phase 2: Authentication (The "Head" Run)
- [ ] **Run Downloader Manually:**
    ```powershell
    # Make sure venv is active
    python scripts/download_rooster.py
    ```
- [ ] **Perform Login:**
    - Chrome window opens -> Enter Microsoft Email/Pass.
    - Approve MFA on phone.
    - **WAIT** for the script to close the browser (do not close it yourself).
- [ ] **Verify Success:**
    - Check that `auth.json` exists in the folder.
    - Check that `Rooster  2026.xlsm` is present and has today's date.

## Phase 3: Automation (Task Scheduler)
- [ ] **Create Task:** Open "Task Scheduler" -> "Create Basic Task".
    - **Name:** `RoosterSync Daily`
    - **Trigger:** Daily at 08:50 (and another for 16:00).
    - **Action:** Start a Program.
    - **Program:** `$HOME\Documents\Rooster\venv\Scripts\python.exe` (Use full path).
    - **Arguments:** `run_workflow.py`
    - **Start in:** `$HOME\Documents\Rooster` (Use full path).
- [ ] **Test Task:** Right-click the new task -> Run.
    - Check `logs/workflow.log` to confirm it worked.

## Phase 4: Git Auth (One-time)
- [ ] **Run Manual Push:**
    ```powershell
    python run_workflow.py
    ```
    - If a GitHub login window pops up, sign in to authorize the PC.

**Done!** The system is now autonomous.
