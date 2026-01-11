# Rooster Calendar Sync: "Subscribe Once" Workflow Plan

## Goal
Establish a zero-friction update system where employees subscribe to a calendar link once. Subsequent updates to the Excel roster automatically propagate to their devices without requiring new links or manual file downloads.

## Architecture

### 1. Hosting Strategy
*   **Platform:** GitHub Pages.
*   **Method:** We will use a dedicated `gh-pages` branch (or `docs/` folder) to host *only* the static `.ics` files. This keeps the repository clean and ensures the Excel source file is never exposed publicly.
*   **URL Structure:** `https://[username].github.io/[repo-name]/calendars/kelvin_murphy.ics`

### 2. Privacy & Security
*   **Source Data:** The `Rooster 2026.xlsm` file will be strictly ignored via `.gitignore` to prevent it from ever being uploaded to GitHub.
*   **Filenames:** Human-readable (e.g., `kelvin_murphy.ics`) for ease of management.
*   **Access:** Links are public but unlisted (security by obscurity).

## Automated Workflow (Windows Lab PC)

We have a master python script `run_workflow.py` that executes these steps:

### Step 1: Ingest (Automated Download)
*   **Tool:** `scripts/download_rooster.py` using **Playwright**.
*   **Action:**
    1.  Launches a Chrome browser.
    2.  Uses saved session cookies (`auth.json`) to bypass Microsoft 2FA.
    3.  Downloads `Rooster 2026.xlsm` from SharePoint.
    4.  Overwrites the local file.
*   **Auth:** Requires a one-time manual login run to generate `auth.json`.

### Step 2: Process (Generate Calendars)
*   **Action:** Runs `scripts/rooster_sync.py`.
*   **Output:** Updates `.ics` files in `output/calendars/`.
*   **Optimization:** The generator uses deterministic IDs to ensure calendar apps see *updates* rather than duplicate events.

### Step 3: Publish (Deploy to Web)
*   **Action:**
    1.  Detects if `.ics` files have changed using Git.
    2.  Commits the changes.
    3.  Pushes specifically to the `gh-pages` branch (or configured publishing source).
*   **Result:** GitHub rebuilds the site, making the new data available at the URLs within ~1-2 minutes.

### Step 4: Distribute (Link Generation)
*   **Action:** The script generates a local text file: `master_links_list.txt` (To be implemented).
*   **Content:** Contains the direct subscription link for every employee found in the Excel file.
    *   *Example:* `Kelvin Murphy: https://kelvinalexander.github.io/Rooster/calendars/kelvin_murphy.ics`

## User Action Required
1.  **Setup:** Follow `WINDOWS_SETUP.md` on the lab PC.
2.  **Monitor:** Check `logs/workflow.log` occasionally.
3.  **Distribute:** Copy links from the master list and send to colleagues once.