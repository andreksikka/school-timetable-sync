# Edupage timetable -> Google Calendar with GitHub Actions

This repo updates a public ICS file from the Järveküla Edupage public timetable and lets Google Calendar subscribe to it.

## What it does
- Fetches the public Edupage timetable JSON
- Filters class `3b`
- Chooses the best lesson when multiple parallel group lessons exist:
  - prefers entries that explicitly mention `3b`
  - otherwise falls back to one deterministic entry
- Generates `3b.ics`
- GitHub Actions updates the file weekly
- Google Calendar subscribes to the raw ICS URL

## Files
- `generate_ics.py` - fetches timetable data and builds `3b.ics`
- `.github/workflows/update.yml` - weekly GitHub Actions workflow

## Setup

### 1. Create a GitHub repo
Create a new public repository, for example:
`jarvekyla-3b-calendar`

### 2. Upload these files
Upload:
- `generate_ics.py`
- `.github/workflows/update.yml`

### 3. Run once manually
In GitHub:
- open the **Actions** tab
- open workflow **Update timetable ICS**
- click **Run workflow**

This should generate and commit `3b.ics`.

### 4. Get the raw ICS URL
After the file exists in your public repo, the raw URL will look like:

`https://raw.githubusercontent.com/YOUR_GITHUB_USER/YOUR_REPO/main/3b.ics`

### 5. Add it to Google Calendar
In Google Calendar on desktop:
- Add calendar
- From URL
- paste the raw ICS URL

Google will subscribe to it. Refresh is delayed and may take a few hours.

## Adjusting dates
In `generate_ics.py`, edit:

```python
DATE_RANGES = [
    ("2026-03-01", "2026-04-12"),
    ("2026-04-20", "2026-06-05"),
]
```

## Adjusting the class
Change:

```python
CLASS_SHORT = "3b"
```

## Notes
- This depends on Edupage keeping the public timetable endpoint structure stable.
- Google Calendar URL subscriptions are not instant.
- The calendar is rebuilt on each run, so the ICS always reflects the current published timetable.
