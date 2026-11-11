#!/usr/bin/env python3
"""
Alert the LINE group about people who have not scanned their fingerprint yet.
Built-in auto run: see SCHEDULE below (day shift & night shift),
every day except Sundays and factory holidays (holidays.txt).

Shift is read from employees.xlsx (column "Group", or "Shift" if there is no Group column):
  D     = always day shift
  A / B = alternate night <-> day every ROTATION_WEEKS weeks
          (anchored at ROTATION_ANCHOR: group A starts a NIGHT block, B is on day)
Saturday working shift (column "Shift": A/B) alternates every 1 week
(Shift B on odd weeks, Shift A on even weeks), and employees who scanned in
on their Saturday day off are highlighted as Off-day OT.
"""
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta, timezone

import requests

URL = "https://webapp.calcomp.co.th/att/getwpr104fig.php"

# Force Thailand Timezone (UTC+7) for GitHub Actions & Local execution
TH_TZ = timezone(timedelta(hours=7))

# All files are looked up next to this script (works no matter where it is started from)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Employee list: employees.xlsx (Excel)
# Columns: Emp.no, Name, Shift (A/B), Group (D / A / B, optional - if missing, Shift is used)
EMP_FILE = os.path.join(BASE_DIR, "employees.xlsx")
HOLIDAY_FILE = os.path.join(BASE_DIR, "holidays.txt")   # one per line: YYYY-MM-DD description

# ==========================================
# ⚙️ LINE TOKEN and GROUP ID
# Set the token as an environment variable (recommended) instead of hard-coding it:
#   Windows (cmd):  setx LINE_TOKEN "your-token"
# ==========================================
LINE_TOKEN = (os.environ.get("LINE_TOKEN") or "").strip()
LINE_GROUP_ID = (os.environ.get("LINE_GROUP_ID") or "C764515bfcbc07701b55517c34943c963").strip()

DUP_SECONDS = 120                      # scans closer than this count as one scan

# Auto run times (minutes since midnight) and how long after the time a run may still start
SCHEDULE = [
    (7 * 60 + 40, "day"),     # 07:40
    (8 * 60 + 10, "day"),     # 08:10
    (19 * 60 + 50, "night"),  # 19:50
    (20 * 60 + 30, "night"),  # 20:30
]
GRACE_MIN = 10

LABEL = {"day": "Day shift", "night": "Night shift"}

# Group A/B change shift (day <-> night) every N weeks
ROTATION_WEEKS = 2
# Reference Monday: from this week on, group A starts a NIGHT block (group B is on day).
# Blocks then alternate every ROTATION_WEEKS weeks (works for dates before the anchor too).
ROTATION_ANCHOR = date(2026, 11, 9)


# ---------- Fetch data ----------
BASE = "https://webapp.calcomp.co.th/att/"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/156.0.0.0 Safari/537.36")

# (connect timeout, read timeout) in seconds - short so a blocked site fails fast
FETCH_TIMEOUT = (5, 10)


def fetch_raw(emp_no):
    r = requests.post(
        URL,
        files={"EMP_NO": (None, emp_no)},
        headers={
            "User-Agent": UA,
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
            "Origin": "https://webapp.calcomp.co.th",
            "Referer": BASE,
            "X-Requested-With": "XMLHttpRequest",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
        },
        timeout=FETCH_TIMEOUT,
    )
    if r.status_code != 200:
        print(f"[HTTP {r.status_code}] {r.text[:500]}", file=sys.stderr)
    r.raise_for_status()
    return r.text


PAT = re.compile(r"(\d{2}/\d{2}/\d{4})\D{0,300}?(\d{2}:\d{2}:\d{2})", re.S)


def parse_scans(raw):
    try:
        stamps = [datetime.strptime(r["STIME"], "%d/%m/%Y %H:%M:%S")
                  for r in json.loads(raw)]
    except (ValueError, KeyError, TypeError):
        stamps = [datetime.strptime(f"{d} {t}", "%d/%m/%Y %H:%M:%S")
                  for d, t in PAT.findall(raw.replace("\\/", "/"))]
    out = []
    for t in sorted(set(stamps)):
        if not out or (t - out[-1]).total_seconds() > DUP_SECONDS:
            out.append(t)
    return out


def get_scans(emp_no, retries=2):
    for i in range(retries):
        try:
            return parse_scans(fetch_raw(emp_no))
        except Exception as e:
            print(f"[WARN] {emp_no} try {i + 1}/{retries}: {e!r}", file=sys.stderr)
            if i == retries - 1:
                return None
            time.sleep(1.5)


# ---------- Shift from the Group column ----------
def minutes(dt):
    return dt.hour * 60 + dt.minute


def group_shift(group, d):
    """
    D   -> always 'day'.
    A/B -> alternate night/day every ROTATION_WEEKS weeks, anchored at ROTATION_ANCHOR
           (anchor block: A = night, B = day).
    Returns 'day' / 'night' / None (unknown or empty group).
    """
    group = (group or "").strip().upper()
    if group == "D":
        return "day"
    if group not in ("A", "B"):
        return None
    monday = d - timedelta(days=d.weekday())
    weeks = (monday - ROTATION_ANCHOR).days // 7
    block = (weeks // ROTATION_WEEKS) % 2          # 0 = anchor block, 1 = swapped block
    a_is_night = (block == 0)
    if group == "A":
        return "night" if a_is_night else "day"
    return "day" if a_is_night else "night"


# ---------- LINE ----------
def push_line(text):
    if not LINE_TOKEN:
        raise RuntimeError("LINE_TOKEN is empty (check the GitHub secret named LINE_TOKEN)")
    if not LINE_GROUP_ID:
        raise RuntimeError("LINE_GROUP_ID is empty")
    for i in range(0, len(text), 4500):
        r = requests.post(
            "https://api.line.me/v2/bot/message/push",
            headers={"Authorization": f"Bearer {LINE_TOKEN}"},
            json={"to": LINE_GROUP_ID, "messages": [{"type": "text", "text": text[i:i + 4500]}]},
            timeout=20,
        )
        if r.status_code != 200:
            # 401 = bad token | 429 = monthly quota used up | 400/403 = bad group id / bot not in group
            raise RuntimeError(f"LINE API {r.status_code}: {r.text}")


# ---------- Employees (Excel) ----------
def _norm(v):
    return re.sub(r"[^a-z0-9]", "", str(v).lower()) if v is not None else ""


def _cell_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _read_rows(path):
    """Return the rows (list of tuples) of the first sheet of an .xlsx file."""
    if not os.path.exists(path):
        raise SystemExit(f"Employee file not found: {path}")
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise SystemExit("openpyxl is not installed. Run:  pip install openpyxl")
    wb = load_workbook(path, data_only=True)
    rows = [tuple(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    wb.close()
    return rows


def load_employees():
    """
    Read employees.xlsx (first sheet).
    Header row: Emp.no, Name, Shift (A/B), Group (D / A / B, optional).
    If there is no Group column, the Shift column (A/B) is used as the group.
    Returns a list of (emp_no, name, sat_shift, group).
    """
    emp_col = name_col = shift_col = group_col = None
    out, seen = [], set()
    for row in _read_rows(EMP_FILE):
        if emp_col is None:
            if not any(c is not None and str(c).strip() for c in row):
                continue
            heads = [_norm(c) for c in row]
            for i, h in enumerate(heads):
                if h in ("empno", "employeeno", "empid", "employeeid", "empcode") and emp_col is None:
                    emp_col = i
                elif h in ("name", "employeename", "empname") and name_col is None:
                    name_col = i
                elif h in ("shift", "satshift", "saturdaygroup", "saturdayshift") and shift_col is None:
                    shift_col = i
                elif h in ("group", "workgroup", "shiftgroup") and group_col is None:
                    group_col = i
            if emp_col is None:
                raise SystemExit(f"{EMP_FILE}: header 'Emp.no' not found in the first row")
            if group_col is None:
                group_col = shift_col          # no Group column -> use Shift (A/B) as the group
            if group_col is None:
                raise SystemExit(f"{EMP_FILE}: header 'Group' or 'Shift' not found in the first row")
            continue
        emp = _cell_text(row[emp_col]) if emp_col < len(row) else ""
        if not emp or emp in seen:
            continue
        seen.add(emp)
        name = _cell_text(row[name_col]) if name_col is not None and name_col < len(row) else ""
        sat_shift = _cell_text(row[shift_col]).upper() if shift_col is not None and shift_col < len(row) else ""
        group = _cell_text(row[group_col]).upper() if group_col < len(row) else ""
        out.append((emp, name, sat_shift, group))
    return out


# ---------- Factory holidays (holidays.txt) ----------
_holiday_cache = {"mtime": None, "dates": set()}


def load_holidays():
    """Dates from holidays.txt ('YYYY-MM-DD description')."""
    try:
        mtime = os.path.getmtime(HOLIDAY_FILE)
    except OSError:
        return set()
    if _holiday_cache["mtime"] == mtime:
        return _holiday_cache["dates"]
    dates = set()
    today_year = datetime.now(TH_TZ).year
    with open(HOLIDAY_FILE, encoding="utf-8-sig", errors="ignore") as f:
        for n, line in enumerate(f, 1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", line)
            if not m:
                print(f"[WARN] holidays.txt line {n} ignored (expected YYYY-MM-DD): {line[:40]}")
                continue
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if y >= 2400:
                y -= 543
            try:
                dt = date(y, mo, d)
            except ValueError:
                print(f"[WARN] holidays.txt line {n} is not a real date: {line[:40]}")
                continue
            if abs(y - today_year) > 5:
                print(f"[WARN] holidays.txt line {n}: year {y} looks wrong: {line[:40]}")
                continue
            dates.add(dt)
    _holiday_cache.update(mtime=mtime, dates=dates)
    return dates


def is_workday(d):
    """No run on Sundays or factory holidays."""
    return d.weekday() != 6 and d not in load_holidays()


def skip_reason(d):
    if d.weekday() == 6:
        return "Sunday"
    if d in load_holidays():
        return "factory holiday"
    return None


def get_working_saturday_shift(d):
    """Returns 'B' for odd weeks and 'A' for even weeks."""
    week_num = d.isocalendar()[1]
    return "A" if week_num % 2 == 0 else "B"


# ---------- Main logic ----------
def run(shift_name, dry):
    """Returns True when everything worked, False when sending/fetching failed."""
    now = datetime.now(TH_TZ).replace(tzinfo=None)
    emps = load_employees()
    print(f"[{now:%Y-%m-%d %H:%M:%S}] loaded {len(emps)} employees, "
          f"LINE_TOKEN set: {bool(LINE_TOKEN)}, group id: {LINE_GROUP_ID[:6]}...")

    is_saturday = (now.weekday() == 5)
    active_sat_shift = get_working_saturday_shift(now.date()) if is_saturday else None

    with ThreadPoolExecutor(max_workers=4) as ex:
        scans_all = list(ex.map(lambda e: get_scans(e[0]), emps))

    all_failed = bool(emps) and all(s is None for s in scans_all)
    if all_failed:
        print("[ERROR] Could not fetch data for ANY employee "
              "(site unreachable / blocked from this machine?)", file=sys.stderr)

    scanned_ok, ot_scanned, missing, unknown, failed, report = [], [], [], [], [], []
    in_shift = 0

    for (emp, name, sat_shift, group), scans in zip(emps, scans_all):
        if scans is None:
            failed.append(emp)
            report.append(f"{emp:<10} {name:<14} could not fetch data")
            continue

        # Shift comes straight from the Group column (D / A / B) - no guessing from scans
        shift = group_shift(group, now.date())
        today = [s for s in scans if s.date() == now.date()]
        last = f"last scan today {today[-1]:%H:%M}" if today else "no scan today"

        # Special Handling for Saturdays (Off-day OT)
        if is_saturday and sat_shift in ("A", "B") and sat_shift != active_sat_shift:
            saturday_scans = [s for s in scans if s.date() == now.date()]

            ot_scans = []
            for s in saturday_scans:
                # Night-shift employee scanning early morning (05:00 - 09:00) = clock-out
                # from Friday night, so ignore it
                if shift == "night" and (5 * 60 <= minutes(s) <= 9 * 60):
                    continue
                ot_scans.append(s)

            if ot_scans:
                scan_time_str = ot_scans[-1].strftime("%H:%M")
                ot_scanned.append((emp, name, scan_time_str))
                report.append(f"{emp:<10} {name:<14} Saturday Off - OT Scanned at {scan_time_str}")
            else:
                report.append(f"{emp:<10} {name:<14} Saturday Off - Skipped")
            continue

        if not shift:
            unknown.append((emp, name))
            report.append(f"{emp:<10} {name:<14} group '{group}' unknown | {last}")
            continue
        if shift != shift_name:
            report.append(f"{emp:<10} {name:<14} {LABEL[shift]} (skipped this round) | {last}")
            continue

        in_shift += 1

        # =========================================================================
        # Check the clock-in scan: day shift vs night shift
        # =========================================================================
        matching_scans = []
        if shift_name == "day":
            for s in today:
                m = minutes(s)
                # Day shift clock-in must be between 05:00 - 08:30
                if 5 * 60 <= m <= 8 * 60 + 30:
                    matching_scans.append(s)
        else:
            # Night shift clock-in must be today from 17:00 onwards
            for s in today:
                if minutes(s) >= 17 * 60:
                    matching_scans.append(s)

        if matching_scans:
            scan_time_str = matching_scans[-1].strftime("%H:%M")
            scanned_ok.append((emp, name, scan_time_str))
            report.append(f"{emp:<10} {name:<14} {LABEL[shift]} scanned at {scan_time_str} | {last}")
        else:
            missing.append((emp, name))
            report.append(f"{emp:<10} {name:<14} {LABEL[shift]} NOT scanned | {last}")

    label = LABEL[shift_name]
    header_extra = f" (Saturday Shift {active_sat_shift})" if is_saturday else ""
    lines = [f"📊 PE ING Status Report: {label}{header_extra} at {now:%H:%M} ({now:%d/%m/%Y})\n"]

    lines.append(f"🟢 Scanned ({len(scanned_ok)} people):")
    if scanned_ok:
        lines += [f" {i}. {e} {n} ({st_time})".rstrip()
                  for i, (e, n, st_time) in enumerate(scanned_ok, 1)]
    else:
        lines.append(" (No employees scanned yet)")

    lines.append("")

    if is_saturday:
        lines.append(f"🟠 day-off OT Scanned ({len(ot_scanned)} people):")
        if ot_scanned:
            lines += [f" {i}. {e} {n} ({st_time})".rstrip()
                      for i, (e, n, st_time) in enumerate(ot_scanned, 1)]
        else:
            lines.append(" (No off-day OT scans)")
        lines.append("")

    lines.append(f"🔴 Not Scanned Yet ({len(missing)} people):")
    if missing:
        lines += [f" {i}. {e} {n}".rstrip() for i, (e, n) in enumerate(missing, 1)]
    else:
        lines.append(" 🎉 Everyone has scanned!")

    if unknown:
        lines.append("\nℹ️ Group Unknown (Not checked): "
                     + ", ".join(f"{e} {n}".strip() for e, n in unknown))
    if failed:
        lines.append("\n⚠️ Fetch Failed: " + ", ".join(failed))

    msg = "\n".join(lines)
    print(f"[{now:%Y-%m-%d %H:%M:%S}] --- {label} status ---")
    print(msg)

    ok = not all_failed
    if not dry:
        try:
            push_line(msg)
            print("--> LINE message sent")
        except Exception as e:
            print(f"--> Failed to send LINE message: {e}", file=sys.stderr)
            ok = False
    return ok


# ---------- Auto mode (daemon loop) ----------
def due_shift(now, done):
    if not is_workday(now.date()):
        return None, None
    m = minutes(now)
    for at, name in SCHEDULE:
        key = (now.date(), at, name)
        if at <= m < at + GRACE_MIN and key not in done:
            return at, name
    return None, None


def start_loop():
    time_strs = [f"{at // 60:02d}:{at % 60:02d}" for at, _ in SCHEDULE]
    print("==================================================")
    print("  ScanAlert Auto Service started")
    print(f"  Runs every day at {', '.join(time_strs)}")
    print("  except Sundays and factory holidays (holidays.txt)")
    print("  (press Ctrl+C to stop)")
    print("==================================================")
    today_th = datetime.now(TH_TZ).date()
    reason = skip_reason(today_th)
    print(f"Today is {'a workday' if not reason else 'NOT a workday (' + reason + ')'}; "
          f"{len(load_holidays())} holiday(s) loaded.")

    done = set()
    while True:
        now = datetime.now(TH_TZ).replace(tzinfo=None)
        done = {k for k in done if k[0] == now.date()}
        at, shift = due_shift(now, done)
        if shift:
            done.add((now.date(), at, shift))
            print(f"It is {now:%H:%M} - checking the {LABEL[shift].lower()}...")
            try:
                run(shift, dry=False)
            except Exception as e:
                print(f"[ERROR] run({shift}) failed: {e!r}")
        time.sleep(20)


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "debug":
        raw = fetch_raw(sys.argv[2])
        print(raw[:1500], "\n--- parsed ---")
        for s in parse_scans(raw)[:20]:
            print(s)
    elif len(sys.argv) >= 2 and sys.argv[1] in ("day", "night"):
        dry = "--dry-run" in sys.argv
        reason = skip_reason(datetime.now(TH_TZ).date())
        if reason and not dry:
            print(f"Today is {reason} - nothing to send.")
        else:
            if reason:
                print(f"[NOTE] Today is {reason} - a real run would be skipped (dry-run continues).")
            # exit code 1 -> GitHub Actions run turns red when fetching or sending fails
            sys.exit(0 if run(sys.argv[1], dry) else 1)
    elif len(sys.argv) >= 2 and sys.argv[1] == "auto":
        start_loop()
    else:
        print("Usage:")
        print("  python app.py debug <emp_no>    # show raw data")
        print("  python app.py day --dry-run     # test the day shift (no LINE message)")
        print("  python app.py day               # send the day shift now")
        print("  python app.py night             # send the night shift now")
        print("  python app.py auto              # stand-by loop: runs automatically at the SCHEDULE times")
