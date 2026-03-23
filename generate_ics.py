#!/usr/bin/env python3
from __future__ import annotations

import json
import hashlib
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.request import urlopen, Request

BASE_URL = "https://jarvekyla.edupage.org"
CLASS_SHORT = "3b"
TIMEZONE = "Europe/Tallinn"
CALENDAR_NAME = "Saskia tunniplaan"
OUTPUT_FILE = "3b.ics"

# Inclusive date ranges for event generation.
DATE_RANGES = [
    ("2026-03-01", "2026-04-12"),
    ("2026-04-20", "2026-06-05"),
]

# Optional extra exclusions. Keep empty if not needed.
EXCLUDED_DATES: set[str] = set()

USER_AGENT = "Mozilla/5.0 (compatible; timetable-sync/1.0)"


def fetch_json(path: str) -> dict:
    req = Request(BASE_URL + path, headers={"User-Agent": USER_AGENT})
    with urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


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
    out = []
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

        # Prefer explicit 3b match in group name or subject text.
        explicit_match = 1 if class_short in group_text or class_short in subject_text else 0

        # Prefer entries with any group label over blank ones.
        has_group = 1 if group_text.strip() else 0

        # Stable fallback to deterministic ordering.
        stable = f"{entry.get('subject_name','')}|{teacher_text}|{group_text}"
        return (explicit_match, has_group, stable)

    return sorted(entries, key=score, reverse=True)[0]


def build_events() -> list[dict]:
    viewer = fetch_json("/timetable/server/ttviewer.js?__func=getTTViewerData")
    tt_num = viewer["r"]["regular"]["default_num"]

    data = fetch_json("/timetable/server/regulartt.js?__func=regularttGetData")
    tables = table_map(data)

    periods = row_index(tables["periods"])
    days = row_index(tables["days"])
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
        "0": 0,
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
    for start_s, end_s in DATE_RANGES:
        start_d = parse_date(start_s)
        end_d = parse_date(end_s)
        for current in daterange(start_d, end_d):
            if current.isoformat() in EXCLUDED_DATES:
                continue
            weekday = current.weekday()  # Monday = 0
            same_day_items = []
            for (day_index, period_id), entry in selected_pattern.items():
                if day_index != weekday:
                    continue
                same_day_items.append((int(period_id), entry))

            for _, entry in sorted(same_day_items, key=lambda x: x[0]):
                start_dt = datetime.strptime(f"{current.isoformat()} {entry['starttime']}", "%Y-%m-%d %H:%M")
                end_dt = datetime.strptime(f"{current.isoformat()} {entry['endtime']}", "%Y-%m-%d %H:%M")
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
                    description_lines.append(f"Grupp: {', '.join([g for g in entry['groupnames'] if g.strip()])}")

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
    output_path.write_text("\r\n".join(fold_ics_line(line) for line in lines) + "\r\n", encoding="utf-8")


def main() -> None:
    output_path = Path(OUTPUT_FILE)
    events = build_events()
    write_ics(events, output_path)
    print(f"Wrote {len(events)} events to {output_path}")


if __name__ == "__main__":
    main()
