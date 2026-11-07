#!/usr/bin/env python3
"""
Alert the LINE group about people who have not scanned their fingerprint yet.
Built-in auto run: 07:55 & 08:30 (day shift), 19:55 & 20:30 (night shift),
every day except Sundays and factory holidays (holidays.txt).
Supports alternating Saturday shifts (Shift B on odd weeks, Shift A on even weeks),
and highlights Off-day OT employees who scanned in on their day off.
"""
import json
import os
import re
import sqlite3
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
EMP_FILE = os.path.join(BASE_DIR, "employees.xlsx")     # columns: Emp.no, Name, Shift (A/B)
HOLIDAY_FILE = os.path.join(BASE_DIR, "holidays.txt")   # one per line: YYYY-MM-DD description
DB_FILE = os.path.join(BASE_DIR, "shifts.db")

# ==========================================
# ⚙️ LINE TOKEN and GROUP ID
# ==========================================
LINE_TOKEN = "8fLcjDj6YKPSWebdzUteioyygkJEf2Qa+1xNRiDwf3jAlUnZGgMSnux2m9HrSnN766GzduxdHfBCpzEX7nLwUcpBbRCX1QD/r8gaXL2ZJX2CJd8NLY09PCI5gfJXwO/JUM1diWYYGfU1xW3rX0LP3AdB04t89/1O/w1cDnyilFU="
LINE_GROUP_ID = "C764515bfcbc07701b55517c34943c963"

# First-scan window used to decide the shift
MORNING_WIN = (5 * 60, 8 * 60 + 30)    # 05:00-08:30 (Day shift)
NIGHT_WIN = (17 * 60, 20 * 60 + 30)    # 17:00-20:30 (Night shift)
DUP_SECONDS = 120                      # scans closer than this count as one scan

# Auto run times (minutes since midnight) and how long after the time a run may still start
SCHEDULE = [
    (7 * 60 + 40, "day"),     # 07:40
    (8 * 60 + 10, "day"),     # 08:10
    (19 * 60 + 50, "night"),  # 19:40
    (20 * 60 + 08, "night"),  # 20:08
]
GRACE_MIN = 10

LABEL = {"day": "Day shift", "night": "Night shift"}

# Employees change shift (day <-> night) every N weeks
ROTATION_WEEKS = 2


# ---------- Fetch data ----------
BASE = "https://webapp.calcomp.co.th/att/"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/156.0.0.0 Safari/537.36")


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
        timeout=20,
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


def get_scans(emp_no, retries=3):
    for i in range(retries):
        try:
            return parse_scans(fetch_raw(emp_no))
        except Exception as e:
            if i == retries - 1:
                print(f"[WARN] {emp_no}: {e}", file=sys.stderr)
                return None
            time.sleep(1.5)


# ---------- Shift detection ----------
def minutes(dt):
    return dt.hour * 60 + dt.minute


def flip_shift(shift):
    return "night" if shift == "day" else "day"


def db():
    con = sqlite3.connect(DB_FILE)
    con.execute("""CREATE TABLE IF NOT EXISTS shift(
        emp_no TEXT, week_start TEXT, shift TEXT,
        PRIMARY KEY(emp_no, week_start))""")
    return con


def week_monday(now):
    return (now - timedelta(days=now.weekday())).replace(
        hour=0, minute=0, second=0, microsecond=0)


def detect_shift_for_week(scans, monday):
    """
    ตรวจหากะของสัปดาห์ที่เริ่มต้นด้วย `monday` จากสแกนแรกสุดในสัปดาห์นั้น
    (ตัดสแกนออกกะดึกของเช้าวันจันทร์ที่เป็นของสัปดาห์ก่อนออก)
    """
    week_start_mon = monday.replace(hour=9, minute=0, second=0)  # เริ่มเช็กหลัง 09:00 น. วันจันทร์
    week_end = monday + timedelta(days=7)

    week_scans = [s for s in scans if week_start_mon <= s < week_end]
    if not week_scans:
        # ถ้าไม่มีสแกนหลัง 09:00 วันจันทร์ ให้ลองดูสแกนช่วง 05:00-08:30 ของวันจันทร์
        mon_morning = [s for s in scans if monday.date() == s.date() and 5 * 60 <= minutes(s) <= 8 * 60 + 30]
        if mon_morning:
            return "day"
        return None

    week_scans.sort()
    first_scan = week_scans[0]
    m = minutes(first_scan)

    if MORNING_WIN[0] <= m <= MORNING_WIN[1]:
        return "day"
    elif NIGHT_WIN[0] <= m <= NIGHT_WIN[1]:
        return "night"

    return None


def get_previous_week_shift(con, emp, cur_monday):
    """
    ค้นหากะล่าสุดที่เคยบันทึกไว้ใน DB ย้อนหลังไปสูงสุด 8 สัปดาห์
    """
    check_monday = cur_monday - timedelta(days=7)
    for _ in range(8):
        wk_str = check_monday.strftime("%Y-%m-%d")
        row = con.execute("SELECT shift FROM shift WHERE emp_no=? AND week_start=?",
                          (emp, wk_str)).fetchone()
        if row and row[0]:
            return row[0], check_monday
        check_monday -= timedelta(days=7)
    return None, None


def ensure_shifts(con, emp, scans, now):
    """
    ระบุกะของพนักงานในสัปดาห์นี้
    """
    cur_monday = week_monday(now)
    wk_str = cur_monday.strftime("%Y-%m-%d")

    # 1. เช็ก DB สัปดาห์ปัจจุบัน
    row = con.execute("SELECT shift FROM shift WHERE emp_no=? AND week_start=?",
                      (emp, wk_str)).fetchone()
    if row and row[0]:
        return row[0], False

    # 2. ตรวจจากสแกนจริงในสัปดาห์นี้
    detected = detect_shift_for_week(scans, cur_monday)
    if detected:
        con.execute("INSERT OR REPLACE INTO shift VALUES(?,?,?)", (emp, wk_str, detected))
        return detected, False

    # 3. คำนวณสลับกะล่วงหน้าถ้ายังไม่มีสแกนในสัปดาห์นี้
    prev_shift, prev_monday = get_previous_week_shift(con, emp, cur_monday)
    if prev_shift and prev_monday:
        weeks_diff = (cur_monday - prev_monday).days // 7
        blocks = weeks_diff // ROTATION_WEEKS
        predicted_shift = prev_shift if (blocks % 2 == 0) else flip_shift(prev_shift)
        return predicted_shift, True

    return None, False


# ---------- LINE ----------
def push_line(text):
    for i in range(0, len(text), 4500):
        r = requests.post(
            "https://api.line.me/v2/bot/message/push",
            headers={"Authorization": f"Bearer {LINE_TOKEN}"},
            json={"to": LINE_GROUP_ID, "messages": [{"type": "text", "text": text[i:i + 4500]}]},
            timeout=20,
        )
        r.raise_for_status()


# ---------- Employees (Excel) ----------
def _norm(v):
    return re.sub(r"[^a-z0-9]", "", str(v).lower()) if v is not None else ""


def _cell_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def load_employees():
    """Read employees.xlsx (first sheet). Header row: Emp.no, Name, Shift (optional A/B)."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise SystemExit("openpyxl is not installed. Run:  pip install openpyxl")
    if not os.path.exists(EMP_FILE):
        raise SystemExit(f"Employee file not found: {EMP_FILE}")

    wb = load_workbook(EMP_FILE, data_only=True)
    ws = wb.active
    emp_col = name_col = shift_col = None
    out, seen = [], set()
    for row in ws.iter_rows(values_only=True):
        if emp_col is None:
            if not any(c is not None and str(c).strip() for c in row):
                continue
            heads = [_norm(c) for c in row]
            for i, h in enumerate(heads):
                if h in ("empno", "employeeno", "empid", "employeeid", "empcode") and emp_col is None:
                    emp_col = i
                elif h in ("name", "employeename", "empname") and name_col is None:
                    name_col = i
                elif h in ("shift", "satshift", "saturdaygroup", "group") and shift_col is None:
                    shift_col = i
            if emp_col is None:
                raise SystemExit(f"{EMP_FILE}: header 'Emp.no' not found in the first row")
            continue
        emp = _cell_text(row[emp_col]) if emp_col < len(row) else ""
        if not emp or emp in seen:
            continue
        seen.add(emp)
        name = _cell_text(row[name_col]) if name_col is not None and name_col < len(row) else ""
        sat_shift = _cell_text(row[shift_col]).upper() if shift_col is not None and shift_col < len(row) else ""
        out.append((emp, name, sat_shift))
    wb.close()
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
    now = datetime.now(TH_TZ).replace(tzinfo=None)
    emps = load_employees()
    con = db()

    is_saturday = (now.weekday() == 5)
    active_sat_shift = get_working_saturday_shift(now.date()) if is_saturday else None

    with ThreadPoolExecutor(max_workers=4) as ex:
        scans_all = list(ex.map(lambda e: get_scans(e[0]), emps))

    scanned_ok, ot_scanned, missing, unknown, failed, report = [], [], [], [], [], []
    in_shift = 0

    for (emp, name, sat_shift), scans in zip(emps, scans_all):
        if scans is None:
            failed.append(emp)
            report.append(f"{emp:<10} {name:<14} could not fetch data")
            continue

        shift, estimated = ensure_shifts(con, emp, scans, now)
        today = [s for s in scans if s.date() == now.date()]
        last = f"last scan today {today[-1]:%H:%M}" if today else "no scan today"

        # Special Handling for Saturdays (Off-day OT)
        if is_saturday and sat_shift in ("A", "B") and sat_shift != active_sat_shift:
            saturday_scans = [s for s in scans if s.date() == now.date()]
            
            ot_scans = []
            for s in saturday_scans:
                # ถ้าพนักงานเป็นกะดึก สแกนช่วงเช้าตรู่ (05:00 - 09:00) คือสแกนออกของคืนวันศุกร์ ตัดทิ้ง
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
            report.append(f"{emp:<10} {name:<14} shift unknown | {last}")
            continue
        if shift != shift_name:
            report.append(f"{emp:<10} {name:<14} {LABEL[shift]} (skipped this round) | {last}")
            continue

        in_shift += 1

        # =========================================================================
        # 🟢 แยกการเช็กสแกนเข้างาน กะเช้า VS กะดึก อย่างชัดเจน
        # =========================================================================
        matching_scans = []
        if shift_name == "day":
            for s in today:
                m = minutes(s)
                # สแกนเข้างานกะเช้าต้องอยู่ช่วง 05:00 - 08:30 น. เท่านั้น
                if 5 * 60 <= m <= 8 * 60 + 30:
                    matching_scans.append(s)
        else:
            # กะดึก (Night Shift): เช็กสแกนตั้งแต่ 17:00 น. ของวันนี้เป็นต้นไป
            today_date = now.date()
            tomorrow_date = today_date + timedelta(days=1)
            for s in scans:
                s_date = s.date()
                m = minutes(s)
                if s_date == today_date and m >= 17 * 60:
                    matching_scans.append(s)
                elif s_date == tomorrow_date and m <= 9 * 60:
                    matching_scans.append(s)

        if matching_scans:
            scan_time_str = matching_scans[-1].strftime("%H:%M")
            scanned_ok.append((emp, name, estimated, scan_time_str))
            report.append(f"{emp:<10} {name:<14} {LABEL[shift]}{' (estimated)' if estimated else ''} scanned at {scan_time_str} | {last}")
        else:
            missing.append((emp, name, estimated))
            report.append(f"{emp:<10} {name:<14} {LABEL[shift]}{' (estimated)' if estimated else ''} NOT scanned | {last}")

    con.commit()

    label = LABEL[shift_name]
    header_extra = f" (Saturday Shift {active_sat_shift})" if is_saturday else ""
    lines = [f"📊 PE ING Status Report: {label}{header_extra} at {now:%H:%M} ({now:%d/%m/%Y})\n"]

    lines.append(f"🟢 Scanned ({len(scanned_ok)} people):")
    if scanned_ok:
        lines += [f" {i}. {e} {n} ({st_time})".rstrip()
                  for i, (e, n, est, st_time) in enumerate(scanned_ok, 1)]
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
        lines += [f" {i}. {e} {n}{' *' if est else ''}".rstrip()
                  for i, (e, n, est) in enumerate(missing, 1)]
        if any(est for _, _, est in missing):
            lines.append(f"\n* Shift estimated based on {ROTATION_WEEKS}-week rotation (no scan recorded this week)")
    else:
        lines.append(" 🎉 Everyone has scanned!")

    if unknown:
        lines.append("\nℹ️ Shift Unknown (Not checked): "
                     + ", ".join(f"{e} {n}".strip() for e, n in unknown))
    if failed:
        lines.append("\n⚠️ Fetch Failed: " + ", ".join(failed))

    msg = "\n".join(lines)
    print(f"[{now:%Y-%m-%d %H:%M:%S}] --- {label} status ---")
    print(msg)
    if not dry:
        try:
            push_line(msg)
            print("--> LINE message sent")
        except Exception as e:
            print(f"--> Failed to send LINE message: {e}")


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
        run(sys.argv[1], "--dry-run" in sys.argv)
    elif len(sys.argv) >= 2 and sys.argv[1] == "auto":
        start_loop()
    else:
        print("Usage:")
        print("  python app.py debug <emp_no>    # show raw data")
        print("  python app.py day --dry-run     # test the day shift (no LINE message)")
        print("  python app.py day               # send the day shift now")
        print("  python app.py night             # send the night shift now")
        print("  python app.py auto              # stand-by loop: runs at 07:55, 08:30, 19:55, 20:30 automatically")
