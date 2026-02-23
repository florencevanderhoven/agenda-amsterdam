import os
import re
import requests

OUTLOOK_ICS_URL = r"https://outlook.office365.com/owa/calendar/89ebda66f073440ba078a3c877245d6f@amsterdam.nl/c00c9d1d76bd4c1b8d3230c0f82f3b74390987971231848833/S-1-8-4130157284-2221344195-4234538342-997209183/reachcalendar.ics"
OUTPUT_FILE = "fixed_calendar.ics"
TZID = "Europe/Amsterdam"

VTIMEZONE_BLOCK = f"""BEGIN:VTIMEZONE
TZID:{TZID}
BEGIN:STANDARD
DTSTART:19701025T030000
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
TZNAME:CET
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:19700329T020000
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
TZNAME:CEST
END:DAYLIGHT
END:VTIMEZONE
"""

def unfold_ics(text: str) -> str:
    return re.sub(r"\r?\n[ \t]", "", text)

def ensure_x_wr_timezone(lines: list[str]) -> list[str]:
    if any(l.strip().startswith("X-WR-TIMEZONE:") for l in lines):
        return lines
    out, inserted = [], False
    for line in lines:
        clean = line.lstrip("\ufeff").strip()
        out.append(line)
        if not inserted and clean.upper() == "BEGIN:VCALENDAR":
            out.append(f"X-WR-TIMEZONE:{TZID}")
            inserted = True
    if not inserted:
        out.insert(0, f"X-WR-TIMEZONE:{TZID}")
    return out

def ensure_vtimezone(lines: list[str]) -> list[str]:
    if any(l.strip() == "BEGIN:VTIMEZONE" for l in lines):
        return lines
    out, inserted = [], False
    for line in lines:
        clean = line.lstrip("\ufeff").strip()
        out.append(line)
        if not inserted and clean.upper() == "BEGIN:VCALENDAR":
            out.append(VTIMEZONE_BLOCK.strip())
            inserted = True
    if not inserted:
        out.insert(0, VTIMEZONE_BLOCK.strip())
    return out

def fix_dt_line(line: str) -> str:
    if not (line.startswith("DTSTART") or line.startswith("DTEND")):
        return line
    if "VALUE=DATE" in line:
        return line

    m = re.match(r"^(DTSTART|DTEND)([^:]*)\:(.*)$", line)
    if not m:
        return line

    key, params, value = m.groups()

    # TZID toevoegen als ontbreekt
    if "TZID=" not in params:
        params = f"{params};TZID={TZID}"

    return f"{key}{params}:{value}"

def main():
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "text/calendar,text/plain,*/*",
    }
    r = requests.get(OUTLOOK_ICS_URL, headers=headers, timeout=30, allow_redirects=True)
    r.raise_for_status()

    raw = unfold_ics(r.text)
    lines = raw.splitlines()

    fixed = [fix_dt_line(l) for l in lines]
    fixed = ensure_x_wr_timezone(fixed)
    fixed = ensure_vtimezone(fixed)

    os.makedirs(".", exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write("\r\n".join(fixed) + "\r\n")

    print("Wrote", OUTPUT_FILE)

if __name__ == "__main__":
    main()
