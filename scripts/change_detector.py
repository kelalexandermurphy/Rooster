"""
Change Detector

Tracks changes to the Excel file and generated ICS files.
Handles detection of new employees, removed employees, and schedule changes.
"""

import hashlib
import json
import os
import shutil
from dataclasses import asdict, dataclass
from datetime import datetime
from typing import Dict, Optional, Set, Tuple


@dataclass
class ChangeReport:
    """Report of changes detected during sync."""

    new_employees: list
    removed_employees: list
    updated_employees: list
    unchanged_employees: list
    excel_file_changed: bool
    timestamp: str

    def has_changes(self) -> bool:
        """Check if any changes were detected."""
        return (
            self.excel_file_changed
            or len(self.new_employees) > 0
            or len(self.removed_employees) > 0
            or len(self.updated_employees) > 0
        )

    def summary(self) -> str:
        """Generate a human-readable summary."""
        lines = [f"Sync Report - {self.timestamp}"]
        lines.append("-" * 40)

        if self.excel_file_changed:
            lines.append("Excel file: CHANGED")
        else:
            lines.append("Excel file: unchanged")

        lines.append(f"New employees: {len(self.new_employees)}")
        for name in self.new_employees:
            lines.append(f"  + {name}")

        lines.append(f"Removed employees: {len(self.removed_employees)}")
        for name in self.removed_employees:
            lines.append(f"  - {name}")

        lines.append(f"Updated: {len(self.updated_employees)}")
        lines.append(f"Unchanged: {len(self.unchanged_employees)}")

        return "\n".join(lines)


class ChangeDetector:
    """Detects changes between sync runs."""

    def __init__(self, state_file: str):
        """
        Initialize the change detector.

        Args:
            state_file: Path to the JSON file storing state between runs
        """
        self.state_file = state_file
        self.state = self._load_state()

    def _load_state(self) -> dict:
        """Load state from file or return empty state."""
        if os.path.exists(self.state_file):
            try:
                with open(self.state_file, "r") as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                pass

        return {"excel_hash": None, "employee_hashes": {}, "last_sync": None}

    def _save_state(self) -> None:
        """Save current state to file."""
        os.makedirs(os.path.dirname(self.state_file), exist_ok=True)
        with open(self.state_file, "w") as f:
            json.dump(self.state, f, indent=2)

    def _file_hash(self, filepath: str) -> str:
        """Calculate MD5 hash of a file."""
        hasher = hashlib.md5()
        with open(filepath, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                hasher.update(chunk)
        return hasher.hexdigest()

    def _content_hash(self, content: bytes) -> str:
        """Calculate MD5 hash of content."""
        return hashlib.md5(content).hexdigest()

    def check_excel_changed(self, excel_path: str) -> bool:
        """
        Check if the Excel file has changed since last sync.

        Args:
            excel_path: Path to the Excel file

        Returns:
            True if file has changed or is new
        """
        if not os.path.exists(excel_path):
            return False

        current_hash = self._file_hash(excel_path)
        previous_hash = self.state.get("excel_hash")

        return current_hash != previous_hash

    def update_excel_hash(self, excel_path: str) -> None:
        """Update the stored Excel file hash."""
        if os.path.exists(excel_path):
            self.state["excel_hash"] = self._file_hash(excel_path)

    def detect_employee_changes(
        self, current_employees: Set[str], existing_files: Set[str]
    ) -> Tuple[Set[str], Set[str]]:
        """
        Detect new and removed employees.

        Args:
            current_employees: Set of employee filenames from current Excel
            existing_files: Set of existing ICS filenames

        Returns:
            Tuple of (new_employees, removed_employees)
        """
        new_employees = current_employees - existing_files
        removed_employees = existing_files - current_employees

        return new_employees, removed_employees

    def check_calendar_changed(self, filename: str, calendar_content: bytes) -> bool:
        """
        Check if a calendar's content has changed.

        Args:
            filename: Employee's ICS filename
            calendar_content: Current calendar content as bytes

        Returns:
            True if content has changed
        """
        current_hash = self._content_hash(calendar_content)
        previous_hash = self.state["employee_hashes"].get(filename)

        return current_hash != previous_hash

    def update_calendar_hash(self, filename: str, calendar_content: bytes) -> None:
        """Update the stored hash for a calendar."""
        self.state["employee_hashes"][filename] = self._content_hash(calendar_content)

    def remove_employee_hash(self, filename: str) -> None:
        """Remove a stored employee hash (when employee is removed)."""
        self.state["employee_hashes"].pop(filename, None)

    def archive_removed_employee(
        self, filename: str, calendars_dir: str, archive_dir: str
    ) -> Optional[str]:
        """
        Move a removed employee's ICS file to archive.

        Args:
            filename: ICS filename
            calendars_dir: Directory with active calendars
            archive_dir: Directory for archived calendars

        Returns:
            Path to archived file, or None if file didn't exist
        """
        source = os.path.join(calendars_dir, filename)

        if not os.path.exists(source):
            return None

        # Create archive directory if needed
        os.makedirs(archive_dir, exist_ok=True)

        # Add timestamp to archived filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        archived_name = f"{filename[:-4]}_{timestamp}.ics"
        dest = os.path.join(archive_dir, archived_name)

        shutil.move(source, dest)

        # Remove from state
        self.remove_employee_hash(filename)

        return dest

    def finalize_sync(self) -> None:
        """Save state after a successful sync."""
        self.state["last_sync"] = datetime.now().isoformat()
        self._save_state()


def get_existing_ics_files(directory: str) -> Set[str]:
    """
    Get set of existing ICS filenames in a directory.

    Args:
        directory: Path to check for ICS files

    Returns:
        Set of filenames (not full paths)
    """
    if not os.path.exists(directory):
        return set()

    return {
        f
        for f in os.listdir(directory)
        if f.endswith(".ics") and os.path.isfile(os.path.join(directory, f))
    }
