"""
NUSched - NU Schedule Desktop Application
==========================================
Fetches the student schedule from NU PowerCampus Self-Service,
displays it in a Tkinter GUI table with selectable rows,
and exports the chosen classes to a standards-compliant ICS file.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import requests
import uuid
import json
from datetime import datetime
import re
import os
import threading


# ═══════════════════════════════════════════════════════════════════════════════
#  Constants
# ═══════════════════════════════════════════════════════════════════════════════

API_URL = "https://register.nu.edu.eg/PowerCampusSelfService/Schedule/Student"

REQUEST_HEADERS = {
    "accept": "application/json",
    "accept-language": "en-US,en-GB;q=0.9,en-ZA;q=0.8,en;q=0.7,ar;q=0.6",
    "cache-control": "max-age=0",
    "content-type": "application/json",
    "priority": "u=1, i",
    "sec-ch-ua": "\"Not(A:Brand\";v=\"8\", \"Chromium\";v=\"144\", \"Google Chrome\";v=\"144\"",
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": "\"Windows\"",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "cookie": (
        "ASP.NET_SessionId=islkk4aslqpogm0ierrjgphz; "
        ".AspNet.Cookies=rpP4AHwFLLNWwmGW6IFd0jYONSmI0j1btMSKy5Yz7Vm6iEtTxbt0WLR2hlkpvKN6xR7vEYOlJTNYmEvlsRbBnuX"
        "_ncGMnd0Rtv8-SwzqvjqcJAMlGUHl7ccyElpZe4pe6RwO-gcEfz_d7VUDZyOUCFOyEVLwATca3PrSYEN9uXKhBOwHKp9cXvWSbVOYywLs"
        "-nDh1Oyk0XRl70nO0xrhg2sxobRtxhadLvMOk6vL5L80ndaiqs8GSD0AGBWWUNdK4-H4Ht_SmEloP-47bLYH_6ofWxQJNeJO5dc-O-d5x"
        "_deyYg5zoHJkoQBeE3uNnGCzdPLg_q8IFn8vqUaQ0-m2EmJQrCOpvJJSKqv9rkR1I-S3esZzUBAZKvDXJi_3LpN4nR2bZQvrZthK97sRMv4"
        "--7VDgSmfgJEmvG1RAlCvyQMTf9r5VEqSK8qqYonUgitH2V66-tse5VL55HjfNvTfT7zAepkboHqmRq9pk7Jk-TQc4XfMHR4Oy4EsAelkX"
        "2XiIVPDxAQ7gNuEb6_kzvxLrVZVRJUIeVlVvIaR-EX9ylK7dHorv-iwZhMZ9o0MZspVg_NwI_C0y1OWbjCOuyGsA"
    ),
    "Referer": "https://register.nu.edu.eg/PowerCampusSelfService/Registration/Schedule",
}

REQUEST_BODY = {
    "personId": 15734,
    "yearTermSession": {
        "year": "2026",
        "term": "SPRG",
        "session": "",
    },
}

# Day name / abbreviation → RRULE BYDAY code
DAY_TO_BYDAY = {
    "sunday": "SU", "monday": "MO", "tuesday": "TU",
    "wednesday": "WE", "thursday": "TH", "friday": "FR", "saturday": "SA",
    "sun": "SU", "mon": "MO", "tue": "TU", "wed": "WE",
    "thu": "TH", "fri": "FR", "sat": "SA",
    "su": "SU", "mo": "MO", "tu": "TU", "we": "WE",
    "th": "TH", "fr": "FR", "sa": "SA",
}

# Semester boundaries — Spring 2026
SEMESTER_START = datetime(2026, 2, 8)   # Sunday, first day of semester week
SEMESTER_UNTIL = "20260521"             # RRULE UNTIL date


# ═══════════════════════════════════════════════════════════════════════════════
#  HTTP Fetch
# ═══════════════════════════════════════════════════════════════════════════════

def fetch_schedule():
    """Send the exact POST request and return parsed JSON data."""
    resp = requests.post(
        API_URL,
        headers=REQUEST_HEADERS,
        json=REQUEST_BODY,
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
        # Build from name parts
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
    section dicts.

    Actual response shape:
        {
          "code": ...,
          "data": {
            "schedule": [
              {
                "sections": [ [], [], [], [ {section}, {section}, ... ] ]
              }
            ]
          }
        }

    The sections array is a list-of-lists — we flatten all non-empty
    sub-lists into one list of section dicts.
    """
    # data → dict with "data" key
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
                # This is one of the sub-lists; collect any dicts inside
                for sec in item:
                    if isinstance(sec, dict):
                        sections.append(sec)
            elif isinstance(item, dict):
                # Directly a section dict
                sections.append(item)

    return sections


def parse_schedule(data):
    """
    Parse the API JSON response into a flat list of course dicts.

    Each dict contains:
        courseName, eventId, eventSubType, section,
        instructors (str), day, dayCode, startTime (tuple), endTime (tuple),
        startTimeStr, endTimeStr, building, room, startDate, endDate
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
        # The API uses "eventName" for the course name
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
            # Schedule info directly on the record
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
    """Try to parse a date string (M/D/YYYY, MM/DD/YYYY, YYYY-MM-DD) to YYYYMMDD."""
    if not date_str:
        return None
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(date_str.strip(), fmt).strftime("%Y%m%d")
        except ValueError:
            continue
    return None


def generate_ics(courses, filepath="schedule_export.ics"):
    """Generate a fully compliant ICS file from the selected course list."""
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
    dt_date_str = SEMESTER_START.strftime("%Y%m%d")

    for course in courses:
        day_code = course.get("dayCode")
        start_time = course.get("startTime")
        end_time = course.get("endTime")

        # Skip entries that lack the minimum info for a valid VEVENT
        if not day_code or not start_time:
            continue

        # Determine UNTIL date — prefer per-course endDate, fallback to constant
        until = SEMESTER_UNTIL
        parsed_end = _parse_date_str(course.get("endDate", ""))
        if parsed_end:
            until = parsed_end

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

        self.courses = []       # parsed course list (mirrors Treeview order)
        self.selected = {}      # Treeview item-id -> bool

        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # Style
        style = ttk.Style()
        style.configure("Treeview", rowheight=26)

        # ── Top frame ──
        top = ttk.Frame(self.root, padding=(12, 10, 12, 4))
        top.pack(fill=tk.X)

        self.fetch_btn = ttk.Button(top, text="Fetch Schedule",
                                    command=self._on_fetch)
        self.fetch_btn.pack(side=tk.LEFT)

        self.status_var = tk.StringVar(value="Press \"Fetch Schedule\" to begin.")
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

        # Scrollbars
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

        # Toggle checkbox on click
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
        """Start fetching the schedule in a background thread."""
        self.fetch_btn.configure(state=tk.DISABLED)
        self.status_var.set("Fetching schedule\u2026")
        threading.Thread(target=self._fetch_worker, daemon=True).start()

    def _fetch_worker(self):
        """Background worker that performs the HTTP request."""
        try:
            data = fetch_schedule()
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

    # ── Table population ──────────────────────────────────────────────────────

    def _populate_table(self, courses):
        """Clear and re-fill the Treeview from the parsed course list."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.selected.clear()

        for c in courses:
            # Build a human-readable time string
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
        """Toggle the check-mark when the user clicks the Sel column."""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)   # "#1", "#2", …
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
        """Handle the Generate ICS File button with confirmation dialog."""
        # Collect only the selected courses
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

        # ── Confirmation popup ──
        confirmed = messagebox.askyesno(
            "Confirm",
            "Are you sure this schedule is correct?\n"
            "The ICS file will now be created.")
        if not confirmed:
            return

        # ── Generate ──
        try:
            generate_ics(selected_courses)
            messagebox.showinfo(
                "Success",
                "ICS file created successfully: schedule_export.ics")
        except Exception as exc:
            messagebox.showerror(
                "Error",
                f"Failed to generate ICS file:\n\n{exc}")


# ═══════════════════════════════════════════════════════════════════════════════
#  Entry Point
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()
