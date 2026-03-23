#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import json
import re
import subprocess
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import urlencode

BASE_URL = "https://jarvekyla.edupage.org"
HOLIDAYS_URL = "https://jarvekyla.edu.ee/vaheajad/"
CLASS_SHORT = "3b"
TIMEZONE = "Europe/Tallinn"
CALENDAR_NAME = "3b tunniplaan"
OUTPUT_FILE = "3b.ics"

# Optional manual exclusions in addition to school holidays.
EXCLUDED_DATES: set[str] = set()

USER_AGENT = "Mozilla/5.0 (compatible; timetable-sync/1.4)"

ESTONIAN_MONTHS = {
    "jaanuar": 1,
    "veebruar": 2,
    "märts": 3,
    "aprill": 4,
    "mai": 5,
    "juuni": 6,
    "juuli": 7,
    "august": 8,
    "september": 9,
    "oktoober": 10,
    "november": 11,
    "detsember": 12,
}


def get_global_range() -> tuple[str, str]:
    today = date.today()

    # Jan-Jun -> current year
    # Jul-Dec -> next year
    year = today.year if today.month <= 6 else today.year + 1

    return (f"{year}-03-01", f"{year}-06-05")


def fetch_text(url: str, referer: str | None = None) -> str:
    cmd = [
        "curl",
        "-4",
        "--silent",
        "--show-error",
        "--location",
        "--connect-timeout",
        "20",
        "--max-time",
        "90",
        "--retry",
        "3",
        "--retry-delay",
        "2",
        "--user-agent",
        USER_AGENT,
        "--header",
        "Accept: text/html,application/json,text/javascript,*/*;q=0.1",
    ]
    if referer:
        cmd.extend(["--referer", referer])
    cmd.append(url)

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(
            f"curl failed for {url}\n"
            f"exit={result.returncode}\n"
            f"stdout={result.stdout[:500]}\n"
            f"stderr={result.stderr[:1500]}"
        )
    return result.stdout


def fetch_json(path: str, params: dict | None = None) -> dict:
    url = BASE_URL + path
    if params:
        sep = "&" if "?" in url else "?"
        url += sep + urlencode(params)

    raw = fetch_text(url, referer=f"{BASE_URL}/timetable/").strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        start = raw.find("(")
        end = raw.rfind(")")
        if start != -1 and end != -1 and end > start:
            inner = raw[start + 1 : end].strip()
            return json.loads(inner)

    raise ValueError(f"Unexpected response format from {url}: {raw[:500]}")


def table_map(data: dict) -> dict[str, dict]:
    tables = data["r"]["dbiAccessorRes"]["tables"]
    return {table["id"]: table for table in tables}


def row_index(table: dict) -> dict[str, dict]:
    return {row["id"]: row for row in table.get("data_rows", [])}


def parse_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def daterange(start: date, end: date):
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def ics_escape(value: str) -> str:
    value = value.replace("\\", "\\\\")
    value = value.replace(";", r"\;").replace(",", r"\,")
    value = value.replace("\n", r"\n")
    return value


def fold_ics_line(line: str) -> str:
    encoded = line.encode("utf-8")
    if len(encoded) <= 75:
        return line

    out: list[str] = []
    current = b""

    for ch in line:
        b = ch.encode("utf-8")
        if len(current) + len(b) > 73:
            out.append(current.decode("utf-8"))
            current = b" " + b
        else:
            current += b

    if current:
        out.append(current.decode("utf-8"))

    return "\r\n".join(out)


def normalize_subject(name: str) -> str:
    replacements = {
        "Ingl k L": "Inglise keel",
        "Inglise keel P": "Inglise keel",
        "Matemaatika P": "Matemaatika",
        "Klassijuhataja": "Klassijuhataja tund",
    }
    return replacements.get(name, name)


def choose_best_lesson(entries: list[dict], class_short: str) -> dict:
    class_short = class_short.lower()

    def score(entry: dict) -> tuple[int, int, str]:
        group_text = " ".join(entry.get("groupnames", [])).lower()
        teacher_text = " ".join(entry.get("teacher_names", [])).lower()
        subject_text = entry.get("subject_name", "").lower()

        # Prefer explicit 3b match in group/subject text
        explicit_match = 1 if class_short in group_text or class_short in subject_text else 0

        # Prefer entries with some group label over empty
        has_group = 1 if group_text.strip() else 0

        stable = f"{entry.get('subject_name', '')}|{teacher_text}|{group_text}"
        return (explicit_match, has_group, stable)

    return sorted(entries, key=score, reverse=True)[0]


def infer_school_year_start(global_start: date) -> int:
    return global_start.year - 1 if global_start.month < 9 else global_start.year


def parse_holiday_date(day_str: str, month_name: str, school_year_start: int) -> date:
    month = ESTONIAN_MONTHS[month_name.lower()]
    year = school_year_start if month >= 9 else school_year_start + 1
    return date(year, month, int(day_str))


def fetch_holiday_ranges(global_start: date, global_end: date) -> list[tuple[date, date]]:
    html = fetch_text(HOLIDAYS_URL, referer=HOLIDAYS_URL)
    school_year_start = infer_school_year_start(global_start)

    section_match = re.search(
        r"Koolivaheajad(.*?)(?:</section>|</main>|</article>|$)",
        html,
        flags=re.IGNORECASE | re.DOTALL,
    )
    source_text = section_match.group(1) if section_match else html

    holiday_pattern = re.compile(
        r"([IVX]+)\s+vaheaeg.*?"
        r"(\d{1,2})\.\s*([A-Za-zÕÄÖÜõäöüŠšŽž]+)"
        r"(?:\s+(\d{4}))?\s*[–-]\s*"
        r"(\d{1,2})\.\s*([A-Za-zÕÄÖÜõäöüŠšŽž]+)"
        r"(?:\s+(\d{4}))?",
        flags=re.IGNORECASE | re.DOTALL,
    )

    ranges: list[tuple[date, date]] = []

    for match in holiday_pattern.finditer(source_text):
        start_day, start_month, start_year_text, end_day, end_month, end_year_text = match.group(2, 3, 4, 5, 6, 7)

        if start_month.lower() not in ESTONIAN_MONTHS or end_month.lower() not in ESTONIAN_MONTHS:
            continue

        if start_year_text:
            start_date = date(int(start_year_text), ESTONIAN_MONTHS[start_month.lower()], int(start_day))
        else:
            start_date = parse_holiday_date(start_day, start_month, school_year_start)

        if end_year_text:
            end_date = date(int(end_year_text), ESTONIAN_MONTHS[end_month.lower()], int(end_day))
        else:
            end_date = parse_holiday_date(end_day, end_month, school_year_start)

        if end_date < start_date:
            continue

        if end_date < global_start or start_date > global_end:
            continue

        clipped_start = max(start_date, global_start)
        clipped_end = min(end_date, global_end)
        ranges.append((clipped_start, clipped_end))

    return ranges


def build_excluded_dates(global_start: date, global_end: date) -> set[str]:
    excluded = set(EXCLUDED_DATES)

    for holiday_start, holiday_end in fetch_holiday_ranges(global_start, global_end):
        for d in daterange(holiday_start, holiday_end):
            excluded.add(d.isoformat())

    return excluded


def build_events() -> list[dict]:
    global_range = get_global_range()
    global_start = parse_date(global_range[0])
    global_end = parse_date(global_range[1])
    excluded_dates = build_excluded_dates(global_start, global_end)

    viewer = fetch_json("/timetable/server/ttviewer.js?__func=getTTViewerData")
    tt_num = viewer["r"]["regular"]["default_num"]

    data = fetch_json(
        "/timetable/server/regulartt.js?__func=regularttGetData",
        params={"tt_num": tt_num},
    )
    tables = table_map(data)

    periods = row_index(tables["periods"])
    classes = row_index(tables["classes"])
    subjects = row_index(tables["subjects"])
    teachers = row_index(tables["teachers"])
    classrooms = row_index(tables["classrooms"])
    lessons = row_index(tables["lessons"])
    cards = tables["cards"]["data_rows"]

    class_row = next((row for row in classes.values() if row.get("short") == CLASS_SHORT), None)
    if not class_row:
        raise RuntimeError(f"Class {CLASS_SHORT!r} not found in timetable data")

    class_id = class_row["id"]

    weekday_map = {
        "0": 0,  # Monday
        "1": 1,
        "2": 2,
        "3": 3,
        "4": 4,
    }

    pattern_entries: dict[tuple[int, str], list[dict]] = {}

    for card in cards:
        lesson = lessons.get(card["lessonid"])
        if not lesson or class_id not in lesson.get("classids", []):
            continue

        period = periods.get(card["period"])
        if not period:
            continue

        subject = subjects.get(lesson["subjectid"], {})
        subject_name = normalize_subject(subject.get("name", "Tund"))

        teacher_names = []
        for teacher_id in lesson.get("teacherids", []):
            teacher = teachers.get(teacher_id)
            if teacher:
                teacher_names.append(teacher.get("short", teacher_id))

        room_names = []
        for room_id in card.get("classroomids", []):
            room = classrooms.get(room_id)
            if room:
                room_names.append(room.get("short", room_id))

        entry = {
            "subject_name": subject_name,
            "subject_short": subject.get("short", ""),
            "teacher_names": teacher_names,
            "room_names": room_names,
            "groupnames": lesson.get("groupnames", []),
            "period": card["period"],
            "starttime": period["starttime"],
            "endtime": period["endtime"],
            "durationperiods": int(lesson.get("durationperiods", 1) or 1),
            "tt_num": tt_num,
        }

        for idx, flag in enumerate(card.get("days", "")):
            if flag == "1" and str(idx) in weekday_map:
                key = (weekday_map[str(idx)], card["period"])
                pattern_entries.setdefault(key, []).append(entry)

    selected_pattern: dict[tuple[int, str], dict] = {}
    for key, entries in pattern_entries.items():
        selected_pattern[key] = choose_best_lesson(entries, CLASS_SHORT)

    all_events = []
    for current in daterange(global_start, global_end):
        if current.isoformat() in excluded_dates:
            continue

        weekday = current.weekday()
        same_day_items = []

        for (day_index, period_id), entry in selected_pattern.items():
            if day_index != weekday:
                continue
            same_day_items.append((int(period_id), entry))

        for _, entry in sorted(same_day_items, key=lambda x: x[0]):
            start_dt = datetime.strptime(
                f"{current.isoformat()} {entry['starttime']}",
                "%Y-%m-%d %H:%M",
            )
            end_dt = datetime.strptime(
                f"{current.isoformat()} {entry['endtime']}",
                "%Y-%m-%d %H:%M",
            )
            duration_min = int((end_dt - start_dt).total_seconds() // 60)

            description_lines = [
                f"Klass: {CLASS_SHORT}",
                f"Kestus: {duration_min} min",
                f"Tunniplaan: {entry['tt_num']}",
            ]
            if entry["teacher_names"]:
                description_lines.append(f"Õpetaja: {', '.join(entry['teacher_names'])}")
            if entry["room_names"]:
                description_lines.append(f"Ruum: {', '.join(entry['room_names'])}")
            if any(name.strip() for name in entry["groupnames"]):
                description_lines.append(
                    f"Grupp: {', '.join([g for g in entry['groupnames'] if g.strip()])}"
                )

            uid_seed = f"{CLASS_SHORT}|{current.isoformat()}|{entry['starttime']}|{entry['subject_name']}"
            uid = hashlib.sha1(uid_seed.encode("utf-8")).hexdigest() + "@jarvekyla-edupage"

            all_events.append(
                {
                    "uid": uid,
                    "summary": entry["subject_name"],
                    "start": start_dt,
                    "end": end_dt,
                    "description": "\n".join(description_lines),
                }
            )

    all_events.sort(key=lambda e: e["start"])
    return all_events


def write_ics(events: list[dict], output_path: Path) -> None:
    now_utc = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//ChatGPT//Edupage GitHub Sync//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:" + ics_escape(CALENDAR_NAME),
        "X-WR-TIMEZONE:" + TIMEZONE,
    ]

    for event in events:
        lines.extend(
            [
                "BEGIN:VEVENT",
                f"UID:{event['uid']}",
                f"DTSTAMP:{now_utc}",
                f"DTSTART;TZID={TIMEZONE}:{event['start'].strftime('%Y%m%dT%H%M%S')}",
                f"DTEND;TZID={TIMEZONE}:{event['end'].strftime('%Y%m%dT%H%M%S')}",
                "SUMMARY:" + ics_escape(event["summary"]),
                "DESCRIPTION:" + ics_escape(event["description"]),
                "END:VEVENT",
            ]
        )

    lines.append("END:VCALENDAR")
    output_path.write_text(
        "\r\n".join(fold_ics_line(line) for line in lines) + "\r\n",
        encoding="utf-8",
    )


def main() -> None:
    output_path = Path(OUTPUT_FILE)
    events = build_events()
    write_ics(events, output_path)
    print(f"Wrote {len(events)} events to {output_path}")
    print(f"Global range used: {get_global_range()[0]} -> {get_global_range()[1]}")


if __name__ == "__main__":
    main()
