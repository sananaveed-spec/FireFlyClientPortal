"""Import Client profiles.docx into data/client-profiles.json."""

from __future__ import annotations

import html
import json
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "FireFly" / "Client profiles.docx"
OUT = ROOT / "data" / "client-profiles.json"

US_STATES = {
    "alabama",
    "alaska",
    "arizona",
    "arkansas",
    "california",
    "colorado",
    "connecticut",
    "delaware",
    "florida",
    "georgia",
    "hawaii",
    "idaho",
    "illinois",
    "indiana",
    "iowa",
    "kansas",
    "kentucky",
    "louisiana",
    "maine",
    "maryland",
    "massachusetts",
    "michigan",
    "minnesota",
    "mississippi",
    "missouri",
    "montana",
    "nebraska",
    "nevada",
    "new hampshire",
    "new jersey",
    "new mexico",
    "new york",
    "north carolina",
    "north dakota",
    "ohio",
    "oklahoma",
    "oregon",
    "pennsylvania",
    "rhode island",
    "south carolina",
    "south dakota",
    "tennessee",
    "texas",
    "utah",
    "vermont",
    "virginia",
    "washington",
    "west virginia",
    "wisconsin",
    "wyoming",
}

CA_PROVINCES = {
    "alberta",
    "british columbia",
    "manitoba",
    "new brunswick",
    "newfoundland and labrador",
    "northwest territories",
    "nova scotia",
    "nunavut",
    "ontario",
    "prince edward island",
    "quebec",
    "saskatchewan",
    "yukon",
}

FIELD_ALIASES = {
    "legal name": "legalName",
    "entity type": "entityType",
    "entity type:": "entityType",
    "country": "country",
    "state": "state",
    "province": "state",
    "location": "location",
    "website": "website",
    "point of contact": "pointOfContact",
    "point of contact email id": "pocEmail",
    "point of contact mobile number": "pocPhone",
    "name of focal person": "pointOfContact",
    "email id": "pocEmail",
    "mobile number": "pocPhone",
    "industry": "industry",  # not stored on client profile
}


def extract_text(docx_path: Path) -> str:
    with zipfile.ZipFile(docx_path) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
    text = re.sub(r"</w:p>", "\n", xml)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    text = text.replace("\xa0", " ").replace("\ufeff", "")
    text = re.sub(r"[ \t]+", " ", text)
    return text


def split_entries(text: str) -> list[tuple[str, str]]:
    """Return (title, body) for each numbered entry."""
    parts = re.split(r"(?m)^\s*(\d+)\.\s+(.+?)\s*$", text)
    # parts: preamble, num, title, body, num, title, body, ...
    entries: list[tuple[str, str]] = []
    i = 1
    while i + 2 < len(parts):
        title = parts[i + 1].strip()
        body = parts[i + 2]
        entries.append((title, body))
        i += 3
    return entries


def parse_fields(body: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    for raw_line in body.splitlines():
        line = raw_line.strip()
        if not line or re.match(r"^\d+\.\s+", line):
            continue
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        key_norm = re.sub(r"\s+", " ", key.strip().lower()).rstrip(":")
        mapped = FIELD_ALIASES.get(key_norm)
        if not mapped:
            # fuzzy: strip trailing punctuation variants
            for alias, dest in FIELD_ALIASES.items():
                if key_norm.startswith(alias.rstrip(":")):
                    mapped = dest
                    break
        if not mapped:
            continue
        value = value.strip(" \t-:;")
        value = re.sub(r"\s+", " ", value)
        if value:
            fields[mapped] = value
    return fields


def resolve_country_state(fields: dict[str, str]) -> tuple[str, str]:
    country_raw = fields.get("country", "").strip()
    state_raw = fields.get("state", "").strip()
    location = fields.get("location", "").strip()

    country = ""
    state = state_raw

    if country_raw:
        cl = country_raw.lower()
        if cl in {"canada", "ca", "canadian"}:
            country = "CA"
        elif cl in {"us", "usa", "united states", "united states of america"}:
            country = "US"
        elif cl in {"not sure", "ns", "unknown", "n/a", "na"}:
            country = "Not Sure"
        else:
            country = (
                "US"
                if "us" in cl
                else ("CA" if "canada" in cl or cl == "ca" else "Not Sure")
            )

    location_unknown = location.lower() in {
        "not sure",
        "n/a",
        "na",
        "unknown",
        "",
    }

    # Explicit Location: Not sure → Country + State = Not Sure
    if location and location.lower() in {"not sure", "n/a", "na", "unknown"}:
        country = "Not Sure"
        state = "Not Sure"
    elif location and not location_unknown:
        loc = location.strip()
        loc_l = loc.lower()
        if loc_l in US_STATES:
            state = loc.title() if loc_l != "california" else "California"
            if not country:
                country = "US"
        elif loc_l in CA_PROVINCES or loc_l in {"toronto"}:
            if loc_l == "toronto":
                state = "Ontario"
            else:
                state = loc.title()
            if not country:
                country = "CA"
        elif not state:
            state = loc
    elif location_unknown and not state:
        state = "Not Sure"

    # Province value "Toronto" → Ontario
    if state.lower() == "toronto":
        state = "Ontario"
        if not country or country == "Not Sure":
            country = "CA"

    if not country:
        if state.lower() in CA_PROVINCES or state.lower() == "ontario":
            country = "CA"
        elif state.lower() in US_STATES:
            country = "US"
        else:
            country = "Not Sure"
            if not state:
                state = "Not Sure"

    if country == "Not Sure" and not state:
        state = "Not Sure"

    if not state:
        state = "Not Sure"

    return country, state


def country_code(country: str) -> str:
    if country == "CA":
        return "CA"
    if country == "Not Sure":
        return "NS"
    return "US"


def format_date(dt: datetime) -> str:
    return f"{dt.year:04d}{dt.month:02d}{dt.day:02d}"


def generate_unique_id(existing: list[str], country: str, dt: datetime) -> str:
    prefix = f"{country_code(country)}-{format_date(dt)}-"
    used = {e.strip().lower() for e in existing if e.strip()}
    max_seq = 0
    for uid in existing:
        u = uid.strip()
        if u.lower().startswith(prefix.lower()):
            try:
                seq = int(u[len(prefix) :])
                max_seq = max(max_seq, seq)
            except ValueError:
                pass
    seq = max_seq + 1
    while True:
        candidate = f"{prefix}{seq:03d}"
        if candidate.lower() not in used:
            return candidate
        seq += 1


def main() -> None:
    text = extract_text(DOCX)
    raw_entries = split_entries(text)
    now = datetime.now(timezone.utc)
    ymd_base = datetime(now.year, now.month, now.day, tzinfo=timezone.utc)

    profiles = []
    unique_ids: list[str] = []
    seen_legal: set[str] = set()

    for index, (title, body) in enumerate(raw_entries):
        fields = parse_fields(body)
        legal = fields.get("legalName", "").strip() or title.strip()
        legal_key = legal.lower()
        # Keep duplicates as distinct rows with suffix if needed
        if legal_key in seen_legal:
            legal = f"{legal} ({index + 1})"
            legal_key = legal.lower()
        seen_legal.add(legal_key)

        country, state = resolve_country_state(fields)
        unique_id = generate_unique_id(unique_ids, country, ymd_base)
        unique_ids.append(unique_id)

        created = ymd_base.replace(microsecond=index * 1000)
        # stagger createdAt slightly for stable ordering
        created_at = created.isoformat().replace("+00:00", "Z")

        entity_type = fields.get("entityType", "").strip()
        entity_lower = entity_type.lower()
        if entity_lower == "business corporation":
            entity_type = "Business Corporation"
        elif entity_lower in {
            "c corp",
            "c corporation",
            "c corpoation",
            "c corpoation",
        } or entity_lower.startswith("c corp"):
            entity_type = "C Corpoation"
        elif entity_lower == "partnership":
            entity_type = "Partnership"
        elif entity_lower in {"s corp", "s corporation"}:
            entity_type = "S Corp"

        profile = {
            "id": f"import-{index + 1}-{unique_id.lower()}",
            "uniqueId": unique_id,
            "legalName": legal,
            "operatingName": title.strip(),
            "website": fields.get("website", ""),
            "industry": fields.get("industry", ""),
            "entityType": entity_type,
            "fiscalYearEnd": "",
            "country": country,
            "state": state,
            "streetAddress": "",
            "ein": "",
            "canadianBusinessNumber": "",
            "gstHst": "",
            "provincialSalesTaxNumbers": "",
            "payrollAccountNumbers": "",
            "ownerName": "",
            "ownerEmail": "",
            "ownerPhone": "",
            "pointOfContact": fields.get("pointOfContact", ""),
            "pocEmail": fields.get("pocEmail", ""),
            "pocPhone": fields.get("pocPhone", ""),
            "createdAt": created_at,
        }
        profiles.append(profile)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(profiles, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {len(profiles)} profiles to {OUT}")
    for p in profiles:
        print(
            f"- {p['uniqueId']} | {p['legalName']} | {p['country']} | {p['state']} | {p['pointOfContact']}"
        )


if __name__ == "__main__":
    main()
