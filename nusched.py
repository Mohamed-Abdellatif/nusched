"""
NUSched - NU Schedule Desktop Application
==========================================
Fetches the student schedule from NU PowerCampus Self-Service,
displays it in a Tkinter GUI table with selectable rows,
and exports the chosen classes to a standards-compliant ICS file.

The user pastes the browser "Copy as fetch" string so any student
can use the tool with their own session.
"""

import subprocess as _sp
import sys as _sys

# Auto-install missing dependencies before anything else
try:
    import requests  # noqa: F401
except ImportError:
    print("Installing required package: requests ...")
    _sp.check_call([_sys.executable, "-m", "pip", "install", "requests"])
    import requests  # noqa: F401

import tkinter as tk
from tkinter import ttk, messagebox
import uuid
import json
from datetime import datetime, timedelta
import re
import os
import sys
import subprocess
import threading


# ═══════════════════════════════════════════════════════════════════════════════
#  Constants
# ═══════════════════════════════════════════════════════════════════════════════

# Day name / abbreviation → RRULE BYDAY code
DAY_TO_BYDAY = {
    "sunday": "SU", "monday": "MO", "tuesday": "TU",
    "wednesday": "WE", "thursday": "TH", "friday": "FR", "saturday": "SA",
    "sun": "SU", "mon": "MO", "tue": "TU", "wed": "WE",
    "thu": "TH", "fri": "FR", "sat": "SA",
    "su": "SU", "mo": "MO", "tu": "TU", "we": "WE",
    "th": "TH", "fr": "FR", "sa": "SA",
}

# BYDAY code → Python weekday number (Monday=0 … Sunday=6)
BYDAY_TO_WEEKDAY = {
    "MO": 0, "TU": 1, "WE": 2, "TH": 3, "FR": 4, "SA": 5, "SU": 6,
}


# ═══════════════════════════════════════════════════════════════════════════════
#  Fetch-command Parser
# ═══════════════════════════════════════════════════════════════════════════════

def parse_fetch_command(text):
    """
    Parse a browser 'Copy as fetch' string into (url, headers_dict, body_dict).

    Expected format (from Chrome DevTools → Network → right-click → Copy as fetch):

        fetch("https://…", {
          "headers": { … },
          "body": "{…}",
          "method": "POST"
        });
    """
    text = text.strip()
    if not text:
        raise ValueError("The pasted text is empty.")

    # ── Extract URL ──
    url_match = re.search(r'fetch\s*\(\s*"([^"]+)"', text)
    if not url_match:
        raise ValueError(
            "Could not find a fetch(\"URL\", …) call.\n"
            "Make sure you right-clicked the request in DevTools "
            "and chose \"Copy as fetch\"."
        )
    url = url_match.group(1)

    # ── Extract the options object (everything between the outer { } ) ──
    rest = text[url_match.end():]
    first_brace = rest.find("{")
    last_brace = rest.rfind("}")
    if first_brace == -1 or last_brace == -1 or last_brace <= first_brace:
        raise ValueError("Could not find the request options object { … }.")

    options_str = rest[first_brace:last_brace + 1]

    try:
        options = json.loads(options_str)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Could not parse options as JSON:\n{exc}") from exc

    headers = options.get("headers", {})
    if not isinstance(headers, dict):
        headers = {}

    body_raw = options.get("body", "{}")
    if isinstance(body_raw, str):
        try:
            body = json.loads(body_raw)
        except json.JSONDecodeError:
            body = {}
    elif isinstance(body_raw, dict):
        body = body_raw
    else:
        body = {}

    return url, headers, body


# ═══════════════════════════════════════════════════════════════════════════════
#  HTTP Fetch
# ═══════════════════════════════════════════════════════════════════════════════

def fetch_schedule(url, headers, body):
    """Send a POST request with the parsed parameters and return JSON data."""
    resp = requests.post(
        url,
        headers=headers,
        json=body,
        timeout=30,
        verify=True,
    )
    resp.raise_for_status()
    data = resp.json()

    # The API returns double-encoded JSON: the response body is a JSON string
    # that itself contains the actual JSON object.  Unwrap it.
    if isinstance(data, str):
        data = json.loads(data)

    return data


# ═══════════════════════════════════════════════════════════════════════════════
#  Response Parsing
# ═══════════════════════════════════════════════════════════════════════════════

def _parse_time(time_str):
    """
    Parse a time string into an (hour, minute) tuple.
    Handles 12-hour ("2:30 PM") and 24-hour ("14:30") formats.
    Returns None on failure.
    """
    if not time_str:
        return None
    time_str = time_str.strip()

    # 12-hour: "2:30 PM", "02:30PM", "12:29 PM"
    m = re.match(r"(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)", time_str)
    if m:
        h, mn, ampm = int(m.group(1)), int(m.group(2)), m.group(3).upper()
        if ampm == "PM" and h != 12:
            h += 12
        elif ampm == "AM" and h == 12:
            h = 0
        return (h, mn)

    # 24-hour: "14:30", "08:30"
    m = re.match(r"(\d{1,2}):(\d{2})", time_str)
    if m:
        return (int(m.group(1)), int(m.group(2)))

    return None


def _normalize_day(day_str):
    """Convert a day name or abbreviation to a BYDAY code (MO, TU, …)."""
    if not day_str:
        return None
    return DAY_TO_BYDAY.get(day_str.strip().lower())


def _extract_instructor_name(instructor):
    """Extract a full name string from an instructor value (str or dict)."""
    if isinstance(instructor, str):
        return instructor.strip()
    if isinstance(instructor, dict):
        full = instructor.get("fullName", "") or ""
        if full.strip():
            return full.strip()
        parts = []
        for key in ("firstName", "first", "middleName", "middle",
                     "lastName", "last", "lastNamePrefix"):
            val = instructor.get(key, "")
            if val and val.strip():
                parts.append(val.strip())
        return " ".join(parts)
    return ""


def _extract_sections(data):
    """
    Navigate the PowerCampus response to extract the flat list of
    registered section dicts.

    Actual response shape:
        {
          "code": …,
          "data": {
            "schedule": [
              {
                "sections": [ [], [], [], [ {section}, … ] ]
              }
            ]
          }
        }
    """
    inner = data
    if isinstance(inner, dict) and "data" in inner:
        inner = inner["data"]
    if not isinstance(inner, dict):
        return []

    schedule_list = inner.get("schedule") or inner.get("studentSchedule") or []
    if not isinstance(schedule_list, list):
        return []

    sections = []
    for schedule_block in schedule_list:
        if not isinstance(schedule_block, dict):
            continue
        raw_sections = schedule_block.get("sections", [])
        if not isinstance(raw_sections, list):
            continue
        for item in raw_sections:
            if isinstance(item, list):
                for sec in item:
                    if isinstance(sec, dict):
                        sections.append(sec)
            elif isinstance(item, dict):
                sections.append(item)

    # Only keep sections the student is actually registered in
    sections = [s for s in sections if s.get("isRegistered", False)]

    return sections


def parse_schedule(data):
    """
    Parse the API JSON response into a flat list of course dicts.
    """
    sections = _extract_sections(data)

    # Fallback: if the above returned nothing, try treating data as a flat list
    if not sections:
        if isinstance(data, list):
            sections = [x for x in data if isinstance(x, dict)]
        elif isinstance(data, dict):
            for v in data.values():
                if isinstance(v, list) and v and isinstance(v[0], dict):
                    sections = v
                    break

    courses = []
    for rec in sections:
        course_name = (rec.get("eventName") or rec.get("courseName")
                       or rec.get("course_name") or rec.get("name") or "")
        event_id = (rec.get("eventId") or rec.get("event_id")
                    or rec.get("id", ""))
        event_sub_type = (rec.get("eventSubType") or rec.get("eventSubtype")
                          or rec.get("subType") or "")
        section = rec.get("section") or rec.get("sectionId") or ""

        # ── Instructors ──
        raw_instr = (rec.get("instructors") or rec.get("instructor")
                     or rec.get("instructorName") or rec.get("instructorNames")
                     or [])
        if isinstance(raw_instr, str):
            instructor_names = [raw_instr]
        elif isinstance(raw_instr, list):
            instructor_names = [_extract_instructor_name(i) for i in raw_instr]
        else:
            instructor_names = [_extract_instructor_name(raw_instr)]
        instructor_names = [n for n in instructor_names if n]
        instructor_str = ", ".join(instructor_names)

        # ── Top-level building / room fallback ──
        bldg_top = (rec.get("buildingName") or rec.get("bldgName")
                     or rec.get("building") or rec.get("orgName") or "")
        room_top = rec.get("roomId") or rec.get("room") or ""

        # ── Semester dates ──
        start_date_str = rec.get("startDate") or ""
        end_date_str = rec.get("endDate") or ""

        # ── Schedule time periods ──
        schedules = (rec.get("schedules")
                     or rec.get("scheduleTimePeriods")
                     or rec.get("scheduleList")
                     or rec.get("schedule")
                     or rec.get("timePeriods")
                     or None)

        if schedules and isinstance(schedules, list):
            for sched in schedules:
                if not isinstance(sched, dict):
                    continue
                day_desc = (sched.get("dayDesc") or sched.get("day")
                            or sched.get("dayName") or "")
                day_code = _normalize_day(day_desc)

                st = _parse_time(sched.get("startTime") or sched.get("start_time") or "")
                et = _parse_time(sched.get("endTime") or sched.get("end_time") or "")

                bldg = (sched.get("bldgName") or sched.get("buildingName")
                        or sched.get("building") or sched.get("orgName")
                        or bldg_top)
                rm = sched.get("roomId") or sched.get("room") or room_top

                st_str = (sched.get("startTime") or sched.get("start_time") or "").strip()
                et_str = (sched.get("endTime") or sched.get("end_time") or "").strip()

                courses.append({
                    "courseName": str(course_name),
                    "eventId": str(event_id),
                    "eventSubType": str(event_sub_type),
                    "section": str(section),
                    "instructors": instructor_str,
                    "day": day_desc,
                    "dayCode": day_code,
                    "startTime": st,
                    "endTime": et,
                    "startTimeStr": st_str,
                    "endTimeStr": et_str,
                    "building": str(bldg),
                    "room": str(rm),
                    "startDate": start_date_str,
                    "endDate": end_date_str,
                })
        else:
            day_desc = (rec.get("dayDesc") or rec.get("day")
                        or rec.get("dayName") or "")
            day_code = _normalize_day(day_desc)

            st = _parse_time(rec.get("startTime") or rec.get("start_time") or "")
            et = _parse_time(rec.get("endTime") or rec.get("end_time") or "")

            st_str = (rec.get("startTime") or rec.get("start_time") or "").strip()
            et_str = (rec.get("endTime") or rec.get("end_time") or "").strip()

            courses.append({
                "courseName": str(course_name),
                "eventId": str(event_id),
                "eventSubType": str(event_sub_type),
                "section": str(section),
                "instructors": instructor_str,
                "day": day_desc,
                "dayCode": day_code,
                "startTime": st,
                "endTime": et,
                "startTimeStr": st_str,
                "endTimeStr": et_str,
                "building": str(bldg_top),
                "room": str(room_top),
                "startDate": start_date_str,
                "endDate": end_date_str,
            })

    return courses


# ═══════════════════════════════════════════════════════════════════════════════
#  ICS Generation
# ═══════════════════════════════════════════════════════════════════════════════

def _format_ics_time(t):
    """Format an (hour, minute) tuple as HHMMSS."""
    if t is None:
        return "000000"
    return f"{t[0]:02d}{t[1]:02d}00"


def _ics_escape(text):
    """Escape special characters for ICS text values."""
    text = text.replace("\\", "\\\\")
    text = text.replace(",", "\\,")
    text = text.replace(";", "\\;")
    text = text.replace("\n", "\\n")
    return text


def _parse_date_str(date_str):
    """Try to parse a date string to a datetime object."""
    if not date_str:
        return None
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None


def _derive_semester_bounds(courses):
    """
    Derive semester start (Sunday of the first week) and end date
    from the courses' startDate / endDate fields.

    Returns (semester_start: datetime, until_str: str "YYYYMMDD").
    Falls back to sensible defaults if dates are missing.
    """
    start_dates = []
    end_dates = []
    for c in courses:
        sd = _parse_date_str(c.get("startDate", ""))
        ed = _parse_date_str(c.get("endDate", ""))
        if sd:
            start_dates.append(sd)
        if ed:
            end_dates.append(ed)

    if start_dates:
        earliest = min(start_dates)
        # Roll back to Sunday of that week (weekday 6 = Sunday)
        days_since_sunday = (earliest.weekday() + 1) % 7  # Mon=0→1, Sun=6→0
        semester_start = earliest - timedelta(days=days_since_sunday)
    else:
        semester_start = datetime(2026, 2, 8)  # fallback

    if end_dates:
        until_str = max(end_dates).strftime("%Y%m%d")
    else:
        until_str = "20260521"  # fallback

    return semester_start, until_str


def generate_ics(courses, filepath="schedule_export.ics"):
    """Generate a fully compliant ICS file from the selected course list."""
    semester_start, default_until = _derive_semester_bounds(courses)

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//NUSched//Schedule Export//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:NU Schedule",
        "X-WR-TIMEZONE:Africa/Cairo",
    ]

    dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    for course in courses:
        day_code = course.get("dayCode")
        start_time = course.get("startTime")
        end_time = course.get("endTime")

        if not day_code or not start_time:
            continue

        # Compute the first occurrence date so DTSTART falls on the
        # correct weekday (avoids phantom events in the first week).
        target_wd = BYDAY_TO_WEEKDAY.get(day_code, semester_start.weekday())
        delta_days = (target_wd - semester_start.weekday()) % 7
        first_date = semester_start + timedelta(days=delta_days)
        dt_date_str = first_date.strftime("%Y%m%d")

        # Per-course end date, else semester default
        until = default_until
        ed = _parse_date_str(course.get("endDate", ""))
        if ed:
            until = ed.strftime("%Y%m%d")

        uid = f"{uuid.uuid4().hex[:8]}-{uuid.uuid4().hex[:8]}@nusched"

        course_name = course.get("courseName", "Unknown Course")
        event_sub_type = course.get("eventSubType", "")
        summary = f"{course_name} - {event_sub_type}" if event_sub_type else course_name

        instructors = course.get("instructors", "")
        description = f"Instructor(s): {_ics_escape(instructors)}" if instructors else ""

        building = course.get("building", "")
        room = course.get("room", "")
        location = f"{building} {room}".strip()

        dtstart = f"{dt_date_str}T{_format_ics_time(start_time)}"
        dtend = f"{dt_date_str}T{_format_ics_time(end_time)}" if end_time else dtstart

        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{uid}")
        lines.append(f"DTSTAMP:{dtstamp}")
        lines.append(f"DTSTART:{dtstart}")
        lines.append(f"DTEND:{dtend}")
        lines.append(f"RRULE:FREQ=WEEKLY;BYDAY={day_code};UNTIL={until}T000000Z")
        lines.append(f"SUMMARY:{summary}")
        if description:
            lines.append(f"DESCRIPTION:{description}")
        if location:
            lines.append(f"LOCATION:{location}")
        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")

    with open(filepath, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")

    return os.path.abspath(filepath)


# ═══════════════════════════════════════════════════════════════════════════════
#  Paste Dialog
# ═══════════════════════════════════════════════════════════════════════════════

class PasteDialog(tk.Toplevel):
    """
    Modal dialog that asks the user to paste the browser
    'Copy as fetch' string.
    """

    INSTRUCTIONS = (
        "1.  Open the NU Self-Service schedule page in Chrome.\n"
        "2.  Press F12 to open DevTools \u2192 Network tab.\n"
        "3.  Reload the schedule page.\n"
        "4.  Find the \"Student\" request (POST).\n"
        "5.  Right-click it \u2192 Copy \u2192 Copy as fetch (Node.js).\n"
        "6.  Paste it below and click OK."
    )

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Paste Request")
        self.geometry("720x440")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()

        self.result = None  # will be (url, headers, body) on success

        self._build_ui()

        # Centre on parent
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(px, 0)}+{max(py, 0)}")

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _build_ui(self):
        # Instructions (top)
        instr_frame = ttk.Frame(self, padding=(12, 10, 12, 2))
        instr_frame.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(instr_frame, text="How to get the request:",
                  font=("", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(instr_frame, text=self.INSTRUCTIONS,
                  justify=tk.LEFT, wraplength=680).pack(anchor=tk.W, pady=(4, 0))

        # Buttons (bottom — pack BEFORE the text area so they always show)
        btn_frame = ttk.Frame(self, padding=(12, 4, 12, 10))
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Button(btn_frame, text="OK", command=self._on_ok).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="Cancel", command=self._on_cancel).pack(
            side=tk.RIGHT, padx=(0, 6))

        # Text area (fills remaining space)
        text_frame = ttk.Frame(self, padding=(12, 6, 12, 6))
        text_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.text = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 9),
                            relief=tk.SUNKEN, borderwidth=1)
        scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL,
                               command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_ok(self):
        raw = self.text.get("1.0", tk.END)
        try:
            self.result = parse_fetch_command(raw)
        except ValueError as exc:
            messagebox.showerror("Parse Error", str(exc), parent=self)
            return
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
#  GUI Application
# ═══════════════════════════════════════════════════════════════════════════════

class ScheduleApp:
    """Main Tkinter GUI for NUSched."""

    CHECK_ON = "\u2713"   # ✓
    CHECK_OFF = ""

    def __init__(self, root):
        self.root = root
        self.root.title("NUSched \u2014 Schedule Exporter")
        self.root.geometry("1150x550")
        self.root.minsize(950, 400)

        self.courses = []
        self.selected = {}

        # Saved request parameters (filled after paste dialog)
        self._req_url = None
        self._req_headers = None
        self._req_body = None

        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        style = ttk.Style()
        style.configure("Treeview", rowheight=26)

        # ── Top frame ──
        top = ttk.Frame(self.root, padding=(12, 10, 12, 4))
        top.pack(fill=tk.X)

        self.fetch_btn = ttk.Button(top, text="Fetch Schedule",
                                    command=self._on_fetch)
        self.fetch_btn.pack(side=tk.LEFT)

        ttk.Button(top, text="Show Tutorial",
                   command=self._on_tutorial).pack(side=tk.LEFT, padx=(8, 0))

        self.status_var = tk.StringVar(
            value="Click \"Fetch Schedule\" to paste your request and load classes.")
        ttk.Label(top, textvariable=self.status_var).pack(side=tk.LEFT, padx=16)

        # ── Table frame ──
        table_frame = ttk.Frame(self.root, padding=(12, 4, 12, 4))
        table_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("sel", "course", "type", "section",
                   "instructors", "day", "time", "building", "room")

        self.tree = ttk.Treeview(table_frame, columns=columns,
                                 show="headings", selectmode="none")

        col_cfg = {
            "sel":         ("Sel",            45,  tk.CENTER),
            "course":      ("Course Name",   190,  tk.W),
            "type":        ("Type",           75,  tk.CENTER),
            "section":     ("Section",        65,  tk.CENTER),
            "instructors": ("Instructor(s)", 210,  tk.W),
            "day":         ("Day",            90,  tk.CENTER),
            "time":        ("Time",          120,  tk.CENTER),
            "building":    ("Building",      170,  tk.W),
            "room":        ("Room",           65,  tk.CENTER),
        }
        for cid, (heading, width, anchor) in col_cfg.items():
            self.tree.heading(cid, text=heading)
            self.tree.column(cid, width=width, anchor=anchor, minwidth=30)

        ysb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL,
                            command=self.tree.yview)
        xsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL,
                            command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)

        # ── Bottom frame ──
        bot = ttk.Frame(self.root, padding=(12, 4, 12, 10))
        bot.pack(fill=tk.X)

        self.sel_all_btn = ttk.Button(bot, text="Select All",
                                      command=self._select_all, state=tk.DISABLED)
        self.sel_all_btn.pack(side=tk.LEFT)

        self.desel_all_btn = ttk.Button(bot, text="Deselect All",
                                        command=self._deselect_all, state=tk.DISABLED)
        self.desel_all_btn.pack(side=tk.LEFT, padx=(6, 0))

        self.gen_btn = ttk.Button(bot, text="Generate ICS File",
                                  command=self._on_generate, state=tk.DISABLED)
        self.gen_btn.pack(side=tk.RIGHT)

    # ── Fetch logic ───────────────────────────────────────────────────────────

    def _on_fetch(self):
        """Open the paste dialog, then fetch in a background thread."""
        dlg = PasteDialog(self.root)
        self.root.wait_window(dlg)

        if dlg.result is None:
            return  # user cancelled

        self._req_url, self._req_headers, self._req_body = dlg.result

        self.fetch_btn.configure(state=tk.DISABLED)
        self.status_var.set("Fetching schedule\u2026")
        threading.Thread(target=self._fetch_worker, daemon=True).start()

    def _fetch_worker(self):
        try:
            data = fetch_schedule(self._req_url, self._req_headers,
                                  self._req_body)
            courses = parse_schedule(data)
            self.root.after(0, self._on_fetch_ok, courses)
        except Exception as exc:
            self.root.after(0, self._on_fetch_err, str(exc))

    def _on_fetch_ok(self, courses):
        self.courses = courses
        self._populate_table(courses)
        self.fetch_btn.configure(state=tk.NORMAL)

        if courses:
            self.status_var.set(f"Loaded {len(courses)} class(es).")
            self.gen_btn.configure(state=tk.NORMAL)
            self.sel_all_btn.configure(state=tk.NORMAL)
            self.desel_all_btn.configure(state=tk.NORMAL)
        else:
            self.status_var.set("No schedule data found.")
            messagebox.showinfo("Info",
                                "No schedule data was found in the API response.")

    def _on_fetch_err(self, msg):
        self.fetch_btn.configure(state=tk.NORMAL)
        self.status_var.set("Fetch failed.")
        messagebox.showerror("Error", f"Failed to fetch schedule:\n\n{msg}")

    # ── Tutorial ──────────────────────────────────────────────────────────────

    def _on_tutorial(self):
        """Open the tutorial video with the system default media player."""
        # Resolve path relative to the script's own directory
        base = os.path.dirname(os.path.abspath(__file__))
        video = os.path.join(base, "assets", "scheduletutorial.mp4")

        if not os.path.isfile(video):
            messagebox.showerror(
                "Not Found",
                f"Tutorial video not found:\n{video}")
            return

        try:
            os.startfile(video)  # Windows
        except AttributeError:
            # macOS / Linux fallback
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.Popen([opener, video])

    # ── Table population ──────────────────────────────────────────────────────

    def _populate_table(self, courses):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.selected.clear()

        for c in courses:
            time_display = ""
            if c["startTime"] and c["endTime"]:
                sh, sm = c["startTime"]
                eh, em = c["endTime"]
                time_display = f"{sh:02d}:{sm:02d} \u2013 {eh:02d}:{em:02d}"
            elif c["startTimeStr"] or c["endTimeStr"]:
                time_display = f"{c['startTimeStr']} \u2013 {c['endTimeStr']}"

            vals = (
                self.CHECK_ON,
                c["courseName"],
                c["eventSubType"],
                c["section"],
                c["instructors"],
                c.get("day", ""),
                time_display,
                c["building"],
                c["room"],
            )
            iid = self.tree.insert("", tk.END, values=vals)
            self.selected[iid] = True

    # ── Checkbox toggle ───────────────────────────────────────────────────────

    def _on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col != "#1":
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return

        current = self.selected.get(iid, True)
        self.selected[iid] = not current

        vals = list(self.tree.item(iid, "values"))
        vals[0] = self.CHECK_ON if not current else self.CHECK_OFF
        self.tree.item(iid, values=vals)

    def _select_all(self):
        for iid in self.tree.get_children():
            self.selected[iid] = True
            vals = list(self.tree.item(iid, "values"))
            vals[0] = self.CHECK_ON
            self.tree.item(iid, values=vals)

    def _deselect_all(self):
        for iid in self.tree.get_children():
            self.selected[iid] = False
            vals = list(self.tree.item(iid, "values"))
            vals[0] = self.CHECK_OFF
            self.tree.item(iid, values=vals)

    # ── ICS generation ────────────────────────────────────────────────────────

    def _on_generate(self):
        children = self.tree.get_children()
        selected_courses = []
        for idx, iid in enumerate(children):
            if self.selected.get(iid, False) and idx < len(self.courses):
                selected_courses.append(self.courses[idx])

        if not selected_courses:
            messagebox.showwarning(
                "No Selection",
                "Please select at least one class to include in the ICS file.")
            return

        confirmed = messagebox.askyesno(
            "Confirm",
            "Are you sure this schedule is correct?\n"
            "The ICS file will now be created.")
        if not confirmed:
            return

        try:
            generate_ics(selected_courses)
            self._show_success_dialog()
        except Exception as exc:
            messagebox.showerror(
                "Error",
                f"Failed to generate ICS file:\n\n{exc}")

    def _show_success_dialog(self):
        """Show success message with Google Calendar import instructions."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Success")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame,
                  text="ICS file created successfully!",
                  font=("", 11, "bold")).pack(anchor=tk.W)

        ttk.Label(frame,
                  text="schedule_export.ics",
                  font=("Consolas", 10)).pack(anchor=tk.W, pady=(2, 12))

        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 12))

        ttk.Label(frame,
                  text="How to import into Google Calendar:",
                  font=("", 10, "bold")).pack(anchor=tk.W)

        steps = (
            "1.  Open Google Calendar (calendar.google.com).\n"
            "2.  Click the gear icon \u2192 Settings.\n"
            "3.  In the left sidebar, click \"Import & export\".\n"
            "4.  Click \"Select file from your computer\".\n"
            "5.  Choose the schedule_export.ics file.\n"
            "6.  Pick which calendar to add it to.\n"
            "7.  Click \"Import\"."
        )
        ttk.Label(frame, text=steps, justify=tk.LEFT,
                  wraplength=420).pack(anchor=tk.W, pady=(6, 12))

        ttk.Label(frame,
                  text="Tip: Create a separate calendar for your schedule\n"
                       "so you can easily delete and re-import if needed.",
                  foreground="gray").pack(anchor=tk.W, pady=(0, 12))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Open File Location",
                   command=lambda: self._open_file_location(dlg)).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="OK",
                   command=dlg.destroy).pack(side=tk.RIGHT)

        # Centre on parent
        dlg.update_idletasks()
        px = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_reqwidth()) // 2
        py = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_reqheight()) // 2
        dlg.geometry(f"+{max(px, 0)}+{max(py, 0)}")

    def _open_file_location(self, parent_dlg=None):
        """Open the folder containing the ICS file and select it."""
        filepath = os.path.abspath("schedule_export.ics")
        try:
            # Windows: open Explorer and highlight the file
            subprocess.Popen(["explorer", "/select,", filepath])
        except Exception:
            try:
                os.startfile(os.path.dirname(filepath))
            except Exception as exc:
                messagebox.showerror("Error", str(exc),
                                     parent=parent_dlg or self.root)


# ═══════════════════════════════════════════════════════════════════════════════
#  Entry Point
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()
