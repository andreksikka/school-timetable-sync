#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import json
import re
import subprocess
from datetime import date, datetime, timedelta
from pathlib import Path

BASE_URL = "https://jarvekyla.edupage.org"
HOLIDAYS_URL = "https://jarvekyla.edu.ee/vaheajad/"
CLASS_SHORT = "3b"
TIMEZONE = "Europe/Tallinn"
CALENDAR_NAME = "3b tunniplaan"
OUTPUT_FILE = "3b.ics"

EXCLUDED_DATES: set[str] = set()

USER_AGENT = "Mozilla/5.0 (compatible; timetable-sync/2.0)"

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


def fetch_post_json(url: str, form_fields: list[tuple[str, str]], referer: str | None = None) -> dict:
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
        "Accept: application/json,text/javascript,*/*;q=0.1",
        "--request",
        "POST",
    ]
    if referer:
        cmd.extend(["--referer", referer])

    for key, value in form_fields:
        cmd.extend(["--data-urlencode", f"{key}={value}"])

    cmd.append(url)

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(
            f"curl POST failed for {url}\n"
            f"exit={result.returncode}\n"
            f"stdout={result.stdout[:500]}\n"
            f"stderr={result.stderr[:1500]}"
        )

    raw = result.stdout.strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        start = raw.find("(")
        end = raw.rfind(")")
        if start != -1 and end != -1 and end > start:
            return json.loads(raw[start + 1:end].strip())

    raise ValueError(f"Unexpected JSON response from {url}: {raw[:500]}")


def parse_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def daterange(start: date, end: date):
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def week_ranges(start: date, end: date) -> list[tuple[date, date]]:
    ranges = []
    current = start
    while current <= end:
        week_end = min(current + timedelta(days=6), end)
        ranges.append((current, week_end))
        current = week_end + timedelta(days=1)
    return ranges


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


def get_needed_part() -> dict:
    return {
        "teachers": ["short", "name", "firstname", "lastname", "callname", "subname", "code", "cb_hidden", "expired"],
        "classes": ["short", "name", "firstname", "lastname", "callname", "subname", "code", "classroomid"],
        "classrooms": ["short", "name", "firstname", "lastname", "callname", "subname", "code"],
        "igroups": ["short", "name", "firstname", "lastname", "callname", "subname", "code"],
        "students": ["short", "name", "firstname", "lastname", "callname", "subname", "code", "classid"],
        "subjects": ["short", "name", "firstname", "lastname", "callname", "subname", "code"],
        "events": ["typ", "name"],
        "event_types": ["name", "icon"],
        "subst_absents": ["date", "absent_typeid", "groupname"],
        "periods": ["short", "name", "firstname", "lastname", "callname", "subname", "code", "period", "starttime", "endtime"],
        "dayparts": ["starttime", "endtime"],
        "dates": ["tt_num", "tt_day"],
    }


def fetch_week_data(week_start: date, week_end: date, school_year_start: int) -> dict:
    url = f"{BASE_URL}/rpr/server/maindbi.js?__func=mainDBIAccessor"
    payload = {
        "__args": [
            None,
            school_year_start,
            {
                "vt_filter": {
                    "datefrom": week_start.isoformat(),
                    "dateto": week_end.isoformat(),
                }
            },
            {
                "op": "fetch",
                "needed_part": get_needed_part(),
                "needed_combos": {},
            },
        ],
        "__gsh": "00000000",
    }

    return fetch_post_json(
        url,
        [
            ("__func", "mainDBIAccessor"),
            ("__args", json.dumps(payload["__args"], ensure_ascii=False, separators=(",", ":"))),
            ("__gsh", payload["__gsh"]),
        ],
        referer=f"{BASE_URL}/timetable/",
    )


def find_table(data: dict, table_name: str) -> list[dict]:
    container = data.get("r", {}).get("dbiAccessorRes", {})
    tables = container.get("tables", [])
    for table in tables:
        if table.get("id") == table_name:
            return table.get("data_rows", [])
    return []


def row_index(rows: list[dict]) -> dict[str, dict]:
    return {row["id"]: row for row in rows if "id" in row}


def choose_best_entry(entries: list[dict], class_short: str) -> dict:
    class_short = class_short.lower()

    def score(entry: dict) -> tuple[int, int, str]:
        group_text = " ".join(entry.get("groupnames", [])).lower()
        teacher_text = " ".join(entry.get("teacher_names", [])).lower()
        subject_text = entry.get("subject_name", "").lower()
        explicit_match = 1 if class_short in group_text or class_short in subject_text else 0
        has_group = 1 if group_text.strip() else 0
        stable = f"{entry.get('subject_name', '')}|{teacher_text}|{group_text}"
        return (explicit_match, has_group, stable)

    return sorted(entries, key=score, reverse=True)[0]


def build_events() -> list[dict]:
    global_range = get_global_range()
    global_start = parse_date(global_range[0])
    global_end = parse_date(global_range[1])
    excluded_dates = build_excluded_dates(global_start, global_end)
    school_year_start = infer_school_year_start(global_start)

    all_events: list[dict] = []

    for week_start, week_end in week_ranges(global_start, global_end):
        week_data = fetch_week_data(week_start, week_end, school_year_start)

        classes = row_index(find_table(week_data, "classes"))
        subjects = row_index(find_table(week_data, "subjects"))
        teachers = row_index(find_table(week_data, "teachers"))
        classrooms = row_index(find_table(week_data, "classrooms"))
        periods = row_index(find_table(week_data, "periods"))
        dates = find_table(week_data, "dates")

        class_row = next((row for row in classes.values() if row.get("short") == CLASS_SHORT), None)
        if not class_row:
            continue

        class_id = class_row["id"]

        tt_num_by_day: dict[str, str] = {}
        for d in dates:
            if "tt_day" in d and "tt_num" in d:
                tt_num_by_day[d["tt_day"]] = str(d["tt_num"])

        cards = find_table(week_data, "cards")
        lessons = row_index(find_table(week_data, "lessons"))

        if not cards or not lessons:
            continue

        by_day_period: dict[tuple[str, str], list[dict]] = {}

        for card in cards:
            lesson = lessons.get(card.get("lessonid", ""))
            if not lesson:
                continue

            if class_id not in lesson.get("classids", []):
                continue

            period = periods.get(card.get("period", ""))
            if not period:
                continue

            subject = subjects.get(lesson.get("subjectid", ""), {})
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

            for tt_day in card.get("dateids", []):
                key = (tt_day, card["period"])
                by_day_period.setdefault(key, []).append(
                    {
                        "subject_name": subject_name,
                        "teacher_names": teacher_names,
                        "room_names": room_names,
                        "groupnames": lesson.get("groupnames", []),
                        "starttime": period["starttime"],
                        "endtime": period["endtime"],
                        "tt_day": tt_day,
                    }
                )

        for (tt_day, period_id), entries in by_day_period.items():
            current = parse_date(tt_day)
            if current < global_start or current > global_end:
                continue
            if current.isoformat() in excluded_dates:
                continue

            entry = choose_best_entry(entries, CLASS_SHORT)

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
            ]

            tt_num = tt_num_by_day.get(tt_day)
            if tt_num:
                description_lines.append(f"Tunniplaan: {tt_num}")

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
        "PRODID:-//ChatGPT//Edupage mainDBI Sync//EN",
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
