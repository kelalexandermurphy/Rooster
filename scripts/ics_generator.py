"""
ICS Calendar Generator

Generates ICS (iCalendar) files for each employee based on their schedule.
Supports timed shifts and all-day events.
"""

import os
from datetime import datetime, timedelta
from typing import Dict, Optional
from uuid import uuid4

import pytz
from excel_parser import Employee, ShiftEntry
from icalendar import Calendar, Event, Timezone, TimezoneDaylight, TimezoneStandard


class ICSGenerator:
    """Generates ICS calendar files for employees."""

    def __init__(self, config: dict):
        """
        Initialize the generator with configuration.

        Args:
            config: Configuration dictionary with shift definitions
        """
        self.config = config
        self.timed_shifts = config.get("timed_shifts", {})
        self.allday_events = config.get("allday_events", {})

        # Calendar settings
        calendar_config = config.get("calendar", {})
        self.timezone_str = calendar_config.get("timezone", "Europe/Amsterdam")
        self.name_prefix = calendar_config.get("name_prefix", "Work Schedule")

        # Load timezone
        self.tz = pytz.timezone(self.timezone_str)

    def generate_calendar(self, employee: Employee) -> Calendar:
        """
        Generate an ICS calendar for an employee.

        Args:
            employee: Employee object with schedule data

        Returns:
            Calendar object
        """
        cal = Calendar()

        # Calendar properties
        cal.add("prodid", "-//Rooster Calendar Sync//rooster-sync//NL")
        cal.add("version", "2.0")
        cal.add("calscale", "GREGORIAN")
        cal.add("method", "PUBLISH")
        cal.add("x-wr-calname", f"{self.name_prefix} - {employee.name}")
        cal.add("x-wr-timezone", self.timezone_str)

        # Add timezone component
        self._add_timezone(cal)

        # Add events for each shift
        for shift in employee.shifts:
            event = self._create_event(shift, employee.name)
            if event:
                cal.add_component(event)

        return cal

    def _add_timezone(self, cal: Calendar) -> None:
        """Add timezone information to the calendar."""
        # For European timezones, add standard/daylight components
        vtimezone = Timezone()
        vtimezone.add("tzid", self.timezone_str)

        # Standard time (winter)
        standard = TimezoneStandard()
        standard.add("tzoffsetfrom", timedelta(hours=2))
        standard.add("tzoffsetto", timedelta(hours=1))
        standard.add("tzname", "CET")
        standard.add("dtstart", datetime(1970, 10, 25, 3, 0, 0))
        standard.add("rrule", {"freq": "yearly", "bymonth": 10, "byday": "-1su"})
        vtimezone.add_component(standard)

        # Daylight time (summer)
        daylight = TimezoneDaylight()
        daylight.add("tzoffsetfrom", timedelta(hours=1))
        daylight.add("tzoffsetto", timedelta(hours=2))
        daylight.add("tzname", "CEST")
        daylight.add("dtstart", datetime(1970, 3, 29, 2, 0, 0))
        daylight.add("rrule", {"freq": "yearly", "bymonth": 3, "byday": "-1su"})
        vtimezone.add_component(daylight)

        cal.add_component(vtimezone)

    def _create_event(self, shift: ShiftEntry, employee_name: str) -> Optional[Event]:
        """
        Create a calendar event for a shift.

        Args:
            shift: ShiftEntry with date and code
            employee_name: Name of the employee (for UID generation)

        Returns:
            Event object or None if shift code is not recognized
        """
        code = shift.code.upper()

        # Check if it's a timed shift
        if code in self.timed_shifts:
            return self._create_timed_event(shift, self.timed_shifts[code], employee_name)

        # Check if it's an all-day event
        if code in self.allday_events:
            return self._create_allday_event(shift, self.allday_events[code], employee_name)

        # Unknown code - skip
        return None

    def _create_timed_event(self, shift: ShiftEntry, shift_config: dict, employee_name: str) -> Event:
        """
        Create a timed event (with specific start/end times).

        Args:
            shift: ShiftEntry with date and code
            shift_config: Configuration for this shift type
            employee_name: Name of the employee (for UID generation)

        Returns:
            Event object
        """
        event = Event()

        # Parse start and end times
        start_time = datetime.strptime(shift_config["start"], "%H:%M").time()
        end_time = datetime.strptime(shift_config["end"], "%H:%M").time()

        # Combine date with times
        start_dt = datetime.combine(shift.date.date(), start_time)
        end_dt = datetime.combine(shift.date.date(), end_time)

        # Handle night shift spanning midnight
        if shift_config.get("spans_midnight", False):
            end_dt += timedelta(days=1)

        # Localize to timezone
        start_dt = self.tz.localize(start_dt)
        end_dt = self.tz.localize(end_dt)

        # Event properties
        event.add("summary", shift_config["name"])
        event.add("dtstart", start_dt)
        event.add("dtend", end_dt)
        event.add("dtstamp", datetime.now(self.tz))
        event.add("uid", self._generate_uid(shift, employee_name))

        # Add description with shift code
        event.add("description", f"Shift code: {shift.code}")

        return event

    def _create_allday_event(self, shift: ShiftEntry, event_config: dict, employee_name: str) -> Event:
        """
        Create an all-day event (no specific times).

        Args:
            shift: ShiftEntry with date and code
            event_config: Configuration for this event type
            employee_name: Name of the employee (for UID generation)

        Returns:
            Event object
        """
        event = Event()

        # All-day events use DATE (not DATETIME)
        event.add("summary", event_config["name"])
        event.add("dtstart", shift.date.date())
        event.add("dtend", shift.date.date() + timedelta(days=1))
        event.add("dtstamp", datetime.now(self.tz))
        event.add("uid", self._generate_uid(shift, employee_name))

        # Add description with shift code
        event.add("description", f"Code: {shift.code}")

        return event

    def _generate_uid(self, shift: ShiftEntry, employee_name: str) -> str:
        """
        Generate a unique identifier for an event.

        This UID is deterministic based on employee, date and code,
        so the same shift always gets the same UID.
        This helps calendar apps update existing events.
        """
        import hashlib
        date_str = shift.date.strftime("%Y%m%d")

        # Create a deterministic hash based on employee and shift details
        raw_id = f"{employee_name}-{date_str}-{shift.code}"
        hash_id = hashlib.md5(raw_id.encode('utf-8')).hexdigest()[:16]

        return f"{date_str}-{shift.code}-{hash_id}@rooster-sync"

    def save_calendar(self, calendar: Calendar, filepath: str) -> None:
        """
        Save a calendar to an ICS file.

        Args:
            calendar: Calendar object to save
            filepath: Path to save the file
        """
        # Ensure directory exists
        os.makedirs(os.path.dirname(filepath), exist_ok=True)

        with open(filepath, "wb") as f:
            f.write(calendar.to_ical())


def generate_ics_files(
    employees: Dict[str, Employee], config: dict, output_dir: str
) -> Dict[str, str]:
    """
    Generate ICS files for all employees.

    Args:
        employees: Dictionary mapping filename to Employee
        config: Configuration dictionary
        output_dir: Directory to save ICS files

    Returns:
        Dictionary mapping filename to full file path
    """
    generator = ICSGenerator(config)
    generated_files = {}

    for filename, employee in employees.items():
        # Generate calendar
        calendar = generator.generate_calendar(employee)

        # Save to file
        filepath = os.path.join(output_dir, filename)
        generator.save_calendar(calendar, filepath)

        generated_files[filename] = filepath

    return generated_files
