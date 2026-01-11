#!/usr/bin/env python3
"""
Rooster Calendar Sync - Main Script

Orchestrates the sync process:
1. Load configuration
2. Parse Excel file
3. Detect changes (new/removed employees)
4. Generate ICS calendar files
5. Archive removed employees
6. Log results

Usage:
    python rooster_sync.py [--config CONFIG_PATH] [--force]

Options:
    --config    Path to config file (default: ../config/config.yaml)
    --force     Force regeneration of all calendars, even if unchanged
"""

import argparse
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

import yaml

# Add scripts directory to path for imports
SCRIPT_DIR = Path(__file__).parent
PROJECT_ROOT = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))

from change_detector import ChangeDetector, ChangeReport, get_existing_ics_files
from excel_parser import Employee, parse_excel
from ics_generator import ICSGenerator


def setup_logging(log_dir: str) -> logging.Logger:
    """Configure logging to both file and console."""
    os.makedirs(log_dir, exist_ok=True)

    log_file = os.path.join(
        log_dir, f"sync_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )

    # Create logger
    logger = logging.getLogger("rooster_sync")
    logger.setLevel(logging.DEBUG)

    # File handler - detailed logging
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_format = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    file_handler.setFormatter(file_format)

    # Console handler - info and above
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_format = logging.Formatter("%(levelname)s: %(message)s")
    console_handler.setFormatter(console_format)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger


def load_config(config_path: str) -> dict:
    """Load configuration from YAML file."""
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def resolve_paths(config: dict, project_root: Path) -> dict:
    """Resolve relative paths in config to absolute paths."""
    config["_excel_path"] = project_root / config["excel_file"]
    config["_output_dir"] = project_root / config["output_dir"]
    config["_archive_dir"] = project_root / config["archive_dir"]
    config["_log_dir"] = project_root / config["log_dir"]
    config["_state_file"] = project_root / "config" / ".sync_state.json"

    return config


def run_sync(config: dict, force: bool = False) -> ChangeReport:
    """
    Run the synchronization process.

    Args:
        config: Configuration dictionary
        force: If True, regenerate all calendars

    Returns:
        ChangeReport with sync results
    """
    logger = logging.getLogger("rooster_sync")

    excel_path = str(config["_excel_path"])
    output_dir = str(config["_output_dir"])
    archive_dir = str(config["_archive_dir"])
    state_file = str(config["_state_file"])

    # Initialize components
    change_detector = ChangeDetector(state_file)
    ics_generator = ICSGenerator(config)

    # Create output directories
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(archive_dir, exist_ok=True)

    # Check if Excel file exists
    if not os.path.exists(excel_path):
        raise FileNotFoundError(
            f"Excel file not found: {excel_path}\n"
            "Please ensure the rooster file is downloaded to the correct location."
        )

    # Check if Excel has changed
    excel_changed = change_detector.check_excel_changed(excel_path)

    if not excel_changed and not force:
        logger.info("Excel file unchanged, checking individual calendars...")
    else:
        if force:
            logger.info("Force mode: regenerating all calendars")
        else:
            logger.info("Excel file changed, processing updates...")

    # Parse Excel file
    logger.info(f"Parsing Excel file: {excel_path}")
    employees = parse_excel(excel_path, config)
    logger.info(f"Found {len(employees)} employees with schedules")

    # Get existing ICS files
    existing_files = get_existing_ics_files(output_dir)
    current_employees = set(employees.keys())

    # Detect new and removed employees
    new_employees, removed_employees = change_detector.detect_employee_changes(
        current_employees, existing_files
    )

    # Track changes for report
    new_list = []
    removed_list = []
    updated_list = []
    unchanged_list = []

    # Archive removed employees
    for filename in removed_employees:
        archived_path = change_detector.archive_removed_employee(
            filename, output_dir, archive_dir
        )
        if archived_path:
            employee_name = filename[:-4].replace("_", " ").title()
            removed_list.append(employee_name)
            logger.info(f"Archived removed employee: {employee_name}")

    # Generate/update calendars for current employees
    for filename, employee in employees.items():
        # Generate calendar
        calendar = ics_generator.generate_calendar(employee)
        calendar_bytes = calendar.to_ical()

        is_new = filename in new_employees
        has_changed = change_detector.check_calendar_changed(filename, calendar_bytes)

        if is_new:
            new_list.append(employee.name)
            logger.info(f"New employee: {employee.name}")

        # Save if new, changed, or forced
        if is_new or has_changed or force:
            filepath = os.path.join(output_dir, filename)
            ics_generator.save_calendar(calendar, filepath)
            change_detector.update_calendar_hash(filename, calendar_bytes)

            if not is_new and has_changed:
                updated_list.append(employee.name)
                logger.debug(f"Updated: {employee.name}")
        else:
            unchanged_list.append(employee.name)

    # Update Excel hash and save state
    change_detector.update_excel_hash(excel_path)
    change_detector.finalize_sync()

    # Create report
    report = ChangeReport(
        new_employees=new_list,
        removed_employees=removed_list,
        updated_employees=updated_list,
        unchanged_employees=unchanged_list,
        excel_file_changed=excel_changed,
        timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )

    return report


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Sync Rooster Excel schedule to ICS calendar files"
    )
    parser.add_argument(
        "--config",
        default=str(PROJECT_ROOT / "config" / "config.yaml"),
        help="Path to configuration file",
    )
    parser.add_argument(
        "--force", action="store_true", help="Force regeneration of all calendars"
    )

    args = parser.parse_args()

    # Load and resolve config
    config = load_config(args.config)
    config = resolve_paths(config, PROJECT_ROOT)

    # Setup logging
    logger = setup_logging(str(config["_log_dir"]))

    logger.info("=" * 50)
    logger.info("Rooster Calendar Sync Started")
    logger.info("=" * 50)

    try:
        # Run sync
        report = run_sync(config, force=args.force)

        # Log summary
        logger.info("")
        logger.info(report.summary())

        if report.has_changes():
            logger.info("")
            logger.info("Sync completed with changes.")
        else:
            logger.info("")
            logger.info("Sync completed. No changes detected.")

        return 0

    except FileNotFoundError as e:
        logger.error(str(e))
        return 1
    except Exception as e:
        logger.exception(f"Sync failed: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
