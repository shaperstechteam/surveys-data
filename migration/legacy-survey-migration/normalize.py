"""Deterministic normalization of merged_properties_final.csv → workbook rows."""
from __future__ import annotations

import html
import json
import posixpath
import re
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any
from urllib.parse import unquote, urlparse

EXPECTED_ROW_COUNT = 7398

PURE_ROLES = {
    "owner": "Owner",
    "caretaker": "Caretaker",
    "care taker": "Caretaker",
    "dealer": "Dealer",
    "manager": "Manager",
    "broker": "Broker",
    "owner's representative": "Owner's Representative",
    "owners representative": "Owner's Representative",
    "representative": "Representative",
    "agent": "Agent",
    "administrator": "Administrator",
    "contractor": "Contractor",
    "friend": "Friend",
    "son": "Son",
    "assistant": "Assistant",
    "power of attorney": "Power of Attorney",
    "managing director": "Managing Director",
    "chief financial officer": "Chief Financial Officer",
    "supervisor": "Supervisor",
}

INVALID_NAMES = {
    "mr.",
    "mr",
    "mr....",
    "mr...",
    "mr .",
    "owner",
    "unknown",
    "a------",
    ".",
    "-",
    "--",
    "n/a",
    "na",
    "nil",
    "contact",
}

CITY_MAP = {
    "jhalum": "Jhelum",
    "jehlum": "Jhelum",
    "muree": "Murree",
    "burawala": "Burewala",
    "huripur": "Haripur",
}

PLACEHOLDER_CITIES = {"", "select", ".", "nil", "n/a"}

PROPERTY_TYPE_MAP = {
    "plot": "LAND_PLOT",
    "office": "OFFICE",
    "house": "RESIDENTIAL",
}

SEED_USERS = {
    "tayyab@shapers.com.pk": "Tayyab",
    "waheed@shapers.com.pk": "Sardar Waheed",
    "ali@shapers.com.pk": "Ali Kazmi",
    "alamdar@shapers.com.pk": "Alamdar Hussain",
    "majid@shapers.com.pk": "Abdul Majid",
    "zaroob@shapers.com.pk": "Zaroob Aslam",
    "abdullah@shapers.com.pk": "Abdullah",
    "hamza@shapers.com.pk": "Hamza Qureshi",
}

SEED_BUSINESSES = [
    ("Khaadi", ["khaadi"]),
    ("Engine", ["engine"]),
    ("J.", ["j.", "junaid jamshed", "junaid jamsheed", "j."]),
    ("Sapphire", ["sapphire"]),
    ("Cambridge", ["cambridge"]),
    ("Ethnic", ["ethnic", "ethnc"]),
    ("Borjan", ["borjan"]),
    ("Cheezious", ["cheezious"]),
    ("Tehzeeb Bakers", ["tehzeeb"]),
    ("Ndure", ["ndure"]),
    ("Bonanza Satrangi", ["bonanza", "satrangi"]),
]

PHONE_RE = re.compile(
    r"(?<!\d)(?:\+?92[\s().-]*3\d{2}(?:[\s().-]*\d){7}|0092[\s().-]*3\d{2}(?:[\s().-]*\d){7}"
    r"|03\d{2}(?:[\s().-]*\d){7}|3\d{2}(?:[\s().-]*\d){7}"
    r"|0(?:21|41|42|51|52|53|55|61|62|64|66|67|68|71|74|81|91|92|93|94|96|99)"
    r"(?:[\s().-]*\d){7,8})(?!\d)"
)

IMAGE_EXT = {"jpg", "jpeg", "png", "gif", "webp", "bmp", "tif", "tiff"}
VIDEO_EXT = {"mp4", "mov", "avi", "webm", "mkv", "3gp", "m4v"}

FLOOR_SLOTS = list(range(1, 12))


def blank(v: Any) -> bool:
    return v is None or str(v).strip() == ""


def s(v: Any) -> str:
    return "" if v is None else str(v).strip()


def dumps(obj: Any) -> str:
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


def parse_json_list(raw: str) -> list:
    raw = s(raw)
    if not raw or raw in ("[]", "null", "{}"):
        return []
    try:
        val = json.loads(raw)
    except json.JSONDecodeError:
        return []
    return val if isinstance(val, list) else []


def parse_urls(raw: str) -> list[str]:
    items = parse_json_list(raw)
    urls: list[str] = []
    if items:
        for x in items:
            if isinstance(x, str) and x.strip().startswith("http"):
                urls.append(x.strip().replace("\\/", "/"))
        return urls
    text = s(raw).replace("\\/", "/")
    if text.startswith("http"):
        return [text]
    return re.findall(r"https?://[^\s\"\\]+", text)


def strip_html(v: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html.unescape(v or ""))).strip()


def phones(cn: str) -> list[str]:
    return [m.group(0) for m in PHONE_RE.finditer(cn or "")]


def normalize_phone(token: str) -> str | None:
    digits = re.sub(r"\D", "", token)
    if digits.startswith("0092"):
        digits = digits[2:]
    if digits.startswith("0") and len(digits) == 11:
        digits = "92" + digits[1:]
    if digits.startswith("3") and len(digits) == 10:
        digits = "92" + digits
    if digits.startswith("92") and 11 <= len(digits) <= 13:
        return "+" + digits
    return None


def name_ok(on: str) -> bool:
    v = s(on)
    if not v:
        return False
    if v.lower() in INVALID_NAMES:
        return False
    if re.fullmatch(r"m+r+\.*", v.lower()):
        return False
    return True


def parse_date(raw: str) -> date | None:
    v = s(raw)
    if not v or v.startswith("0000-00-00"):
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(v[:19] if fmt.endswith("%S") else v[:10], fmt).date()
        except ValueError:
            continue
    return None


def parse_decimal(raw: str) -> Decimal | None:
    v = s(raw).replace(",", "")
    if not v:
        return None
    try:
        return Decimal(v)
    except InvalidOperation:
        return None


def parse_area_native(raw: str) -> tuple[Decimal | None, str | None, str | None]:
    """Return (value, unit, warning). Never guess the unit."""
    v = s(raw)
    if not v:
        return None, None, None
    m = re.match(
        r"^([\d]+(?:\.\d+)?)\s*(kanal|kanals|marla|marlas|acre|acres|sq\.?\s*y(?:ar)?ds?|gazz?|sq\.?\s*f(?:ee)?t|sft|sqft)\s*\.?$",
        v,
        re.I,
    )
    if not m:
        return None, None, "AMBIGUOUS_AREA"
    num = Decimal(m.group(1))
    unit_raw = m.group(2).lower()
    if unit_raw.startswith("kanal"):
        unit = "Kanal"
    elif unit_raw.startswith("marla"):
        unit = "Marla"
    elif unit_raw.startswith("acre"):
        unit = "Acre"
    elif "y" in unit_raw or "gaz" in unit_raw:
        unit = "Sq Yd"
    else:
        unit = "Sq Ft"
    return num, unit, None


def parse_sqft(raw: str) -> tuple[Decimal | None, str | None]:
    v = s(raw)
    if not v:
        return None, None
    approx = bool(re.search(r"approx", v, re.I))
    m = re.match(
        r"^([\d,]+(?:\.\d+)?)\s*(?:sq\.?\s*f(?:ee)?t|sft|sqft)?\s*(?:\(?approx\.?\)?)?\.?$",
        v,
        re.I,
    )
    if not m:
        return None, "AMBIGUOUS_AREA"
    num = Decimal(m.group(1).replace(",", ""))
    return num, ("AREA_APPROX" if approx else None)


def parse_int_field(raw: str) -> int | None:
    v = s(raw)
    if not v:
        return None
    try:
        d = Decimal(v)
    except InvalidOperation:
        return None
    if d != d.to_integral_value():
        return None
    return int(d)


def normalize_city(raw: str) -> tuple[str | None, str | None]:
    v = s(raw)
    if v.lower() in PLACEHOLDER_CITIES:
        return None, "PLACEHOLDER_CITY"
    mapped = CITY_MAP.get(v.lower())
    if mapped:
        return mapped, "CITY_SPELLING"
    return v, None


AREA_PLACEHOLDERS = {"", "select", ".", "nil", "n/a", "area not listed"}


def usable_area_value(raw: str) -> str:
    v = s(raw)
    return "" if v.lower() in AREA_PLACEHOLDERS else v


def pick_area(row: dict, city: str | None) -> tuple[str, str]:
    """One area string. Returns (value, sourceColumn)."""
    city_l = (city or "").lower()
    if city_l == "lahore":
        order = ["Areas_In_Lahore", "Market_Area_114", "Map_Search_Address"]
    elif city_l == "islamabad":
        order = ["Areas_In_Islamabad", "Market_Area_114", "Map_Search_Address"]
    elif city_l == "rawalpindi":
        order = ["Areas_In_Rawalpindi", "Market_Area_114", "Map_Search_Address"]
    elif city_l == "karachi":
        order = ["Areas_In_Karachi", "Market_Area_117", "Map_Search_Address"]
    else:
        order = ["Areas_In_Karachi", "Market_Area_114", "Market_Area_117", "Map_Search_Address"]
    for col in order:
        val = usable_area_value(row.get(col) or "")
        if val:
            return val, col
    return "", ""


def parse_asking_amount(raw: str) -> tuple[Decimal | None, str | None]:
    v = s(raw)
    if not v:
        return None, None
    cleaned = re.sub(r"(?i)\b(pkr|rs\.?|rupees)\b", "", v).replace(",", "").strip()
    if re.fullmatch(r"-?\d+(?:\.\d+)?", cleaned):
        return Decimal(cleaned), None
    return None, "UNPARSEABLE_ASKING_AMOUNT"


def parse_datetime_iso(raw: str) -> str:
    d = parse_date(raw)
    return d.isoformat() if d else ""


def coords(lat_raw: str, lng_raw: str) -> tuple[str | None, list[str]]:
    warnings: list[str] = []
    if blank(lat_raw) and blank(lng_raw):
        warnings.append("MISSING_COORDINATES")
        return None, warnings
    if blank(lat_raw) or blank(lng_raw):
        warnings.append("INVALID_COORDINATES")
        return None, warnings
    try:
        lat = float(lat_raw)
        lng = float(lng_raw)
    except ValueError:
        warnings.append("INVALID_COORDINATES")
        return None, warnings
    if not (-90 <= lat <= 90 and -180 <= lng <= 180):
        warnings.append("INVALID_COORDINATES")
        return None, warnings
    if not (23 <= lat <= 38 and 60 <= lng <= 78):
        warnings.append("COORD_OUTSIDE_PAKISTAN")
    return dumps({"latitude": lat, "longitude": lng}), warnings


def match_business(name: str) -> str | None:
    n = s(name).lower()
    if not n:
        return None
    for official, aliases in SEED_BUSINESSES:
        for a in aliases:
            if a == n or a in n or n in a:
                return official
    return None


def classify_contact(row: dict) -> dict:
    on = s(row.get("Owner_Name"))
    cp = s(row.get("Contact_Person"))
    cn = s(row.get("Contact_Number"))
    bd = s(row.get("Business_Developer"))
    cp_l = cp.lower()
    tokens = phones(cn)
    role = PURE_ROLES.get(cp_l)
    internal = bool(
        bd
        and "@" in bd
        and cp
        and (
            cp_l in bd.lower()
            or bd.split("@")[0].replace(".", " ") in cp_l
            or cp_l.replace(" ", "") in bd.lower().replace(".", "")
        )
    )
    mixed = bool(re.search(r"[()/,-]", cp)) and not role
    person_like = (
        not role
        and not internal
        and bool(re.search(r"[A-Za-z]{3,}", cp))
        and cp_l not in {".", "nil", "select"}
    )
    placeholder_cp = cp_l in {".", "nil", "select", "`"} or bool(re.fullmatch(r"[\d* -]+", cp))

    if role and name_ok(on) and len(tokens) >= 2:
        status, action = "REVIEW_MULTI_PHONE", "HOLD"
    elif role and name_ok(on) and len(tokens) == 1:
        status, action = "VALID", "MIGRATE"
    elif role and name_ok(on):
        status, action = "VALID_NO_PHONE", "MIGRATE_PHONE_NULL"
    elif role and not name_ok(on):
        status, action = "INVALID_NAME", "HOLD"
    elif internal or mixed or person_like or placeholder_cp or not role:
        status, action = "REVIEW_CP_SEMANTICS", "HOLD"
    else:
        status, action = "REVIEW_CP_SEMANTICS", "HOLD"

    phone_norm = normalize_phone(tokens[0]) if len(tokens) == 1 else None
    note = {
        "REVIEW_MULTI_PHONE": "Multiple strict phones in one cell; do not split",
        "INVALID_NAME": "Required name is title/placeholder; do not fabricate",
        "REVIEW_CP_SEMANTICS": (
            "Internal BD identity suspected"
            if internal
            else "cp is not a pure role word"
        ),
        "VALID_NO_PHONE": "Safe contact; phone left null",
        "VALID": "Safe contact",
    }[status]

    return {
        "legacySurveyId": int(row["id"]),
        "legacyContactSource": "property-contact:on-cp-cn",
        "legacyContactIndex": 0,
        "name": on if name_ok(on) and action.startswith("MIGRATE") else "",
        "role": role or "",
        "phone": phone_norm or "",
        "email": "",
        "legacyName": on,
        "legacyRole": cp,
        "legacyPhone": cn,
        "validationStatus": status,
        "migrationStatus": action,
        "migrationNotes": note,
        "migrationAction": action,
        "notes": note,
    }


def user_match(bd: str) -> tuple[str, str]:
    email = s(bd).lower()
    if not email or email == "select" or "@" not in email:
        return "", "UNMATCHED"
    if email in SEED_USERS:
        return email, "MATCH_SEED_CANDIDATE"
    return email, "UNMATCHED"


def build_units(row: dict) -> tuple[str | None, list[str]]:
    warnings: list[str] = []
    units: list[dict] = []
    for i in FLOOR_SLOTS:
        ftype = s(row.get(f"Floor_{i}_Type"))
        farea = s(row.get(f"Floor_{i}_Area"))
        fsize = s(row.get(f"Floor_{i}_Size"))
        if not ftype and not farea and not fsize:
            continue
        unit: dict[str, Any] = {
            "number": ftype or f"Floor {i}",
            "floor": ftype or f"Floor {i}",
        }
        sqft, warn = parse_sqft(farea) if farea else (None, None)
        if sqft is not None:
            unit["squareFeet"] = float(sqft)
        elif farea:
            warnings.append(f"AMBIGUOUS_FLOOR_AREA:{i}")
        if fsize:
            unit["dimensions"] = fsize
        units.append(unit)

    if not units:
        labels = [str(x) for x in parse_json_list(row.get("Select_Floors") or "") if str(x).strip()]
        if labels:
            units = [{"number": lab, "floor": lab} for lab in labels]
            warnings.append("UNITS_FROM_SELECT_FLOORS_LABELS_ONLY")

    if not units:
        return None, warnings
    return dumps(units), warnings


def build_market_details(row: dict) -> tuple[str | None, list[str], list[str]]:
    """Returns (json, warnings, competitor names)."""
    names: list[str] = []
    for brand in parse_json_list(row.get("Select_Brands") or ""):
        if isinstance(brand, str) and s(brand):
            names.append(s(brand))
    for extra in (s(row.get("Other_Brands")),):
        if extra:
            names.append(extra)
    earlier = s(row.get("Earlier_Brands"))
    if earlier:
        names.append(earlier)

    # unique preserve order
    seen = set()
    unique = []
    for n in names:
        key = n.lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(n)
    if not unique:
        return None, [], []
    return dumps({"nearbyCompetitors": unique}), [], unique


def filename_of(url: str) -> tuple[str, str]:
    path = unquote(urlparse(url).path)
    name = posixpath.basename(path) or "file"
    ext = name.rsplit(".", 1)[1].lower() if "." in name else ""
    return name, ext


def classify_attachment_type(ext: str, source: str) -> str:
    if ext in IMAGE_EXT:
        return "IMAGE"
    if ext in VIDEO_EXT:
        return "VIDEO"
    if source == "Videos":
        return "VIDEO"
    if source == "Pictures":
        return "IMAGE"
    return "OTHER"


def attachment_status(url: str, head: dict[str, dict], atype: str) -> tuple[str, str, str]:
    """sourceExists, migrationStatus, notes"""
    info = head.get(url) or head.get(url.replace("https://", "http://")) or head.get(
        url.replace("http://", "https://")
    )
    placeholder = url.rstrip("/").endswith("/1")
    if placeholder:
        return "N", "HOLD", "Placeholder Gravity Forms path ending in /1"
    if atype == "OTHER":
        exists = "UNKNOWN"
        if info:
            final = info.get("final_url") or ""
            exists = "Y" if info.get("status") == "200" and "wp-login.php" not in final else "N"
        return exists, "HOLD", "AttachmentRecord only allows IMAGE|VIDEO"
    if not info:
        return "UNKNOWN", "REVIEW", "URL not in prior HEAD scan"
    final = info.get("final_url") or ""
    if info.get("status") == "200" and "wp-login.php" not in final:
        return "Y", "READY", ""
    return "N", "HOLD", "HEAD indicated missing or WordPress login redirect"


def load_head_scan(path) -> dict[str, dict]:
    import csv

    out = {}
    if not path or not path.exists():
        return out
    with path.open(encoding="utf-8", newline="") as fh:
        for row in csv.DictReader(fh):
            out[row["url"]] = row
    return out


def survey_status(warnings: list[str], review_reasons: list[str]) -> str:
    high = {
        "INVALID_COORDINATES",
        "INVALID_DATE",
        "UNKNOWN_PROPERTY_TYPE",
        "PLACEHOLDER_CITY",
        "MISSING_SURVEY_DATE",
    }
    if any(r.startswith("UNKNOWN_PROPERTY_TYPE") or r == "UNKNOWN_PROPERTY_TYPE" for r in review_reasons):
        return "REVIEW"
    if any(w in high or w.startswith("AMBIGUOUS") for w in warnings) or review_reasons:
        return "READY_WITH_WARNINGS" if not review_reasons else "REVIEW"
    if warnings:
        return "READY_WITH_WARNINGS"
    return "READY"


def normalize_survey(row: dict) -> tuple[dict, list[dict], list[dict]]:
    """Return (survey_row, issues, decisions)."""
    warnings: list[str] = []
    reviews: list[str] = []
    issues: list[dict] = []
    decisions: list[dict] = []
    lid = int(row["id"])

    def issue(field, raw, itype, severity, auto, norm, action, notes=""):
        issues.append(
            {
                "legacySurveyId": lid,
                "field": field,
                "rawValue": raw if raw is None else str(raw)[:32000],
                "issueType": itype,
                "severity": severity,
                "automaticFixPossible": auto,
                "normalizedValue": norm or "",
                "migrationAction": action,
                "notes": notes,
            }
        )

    city, city_warn = normalize_city(row.get("City"))
    if city_warn:
        warnings.append(city_warn)
        issue("City", row.get("City"), city_warn, "MEDIUM", "Y" if city else "N", city or "", "WARN")

    area, area_source = pick_area(row, city)

    loc, coord_warns = coords(row.get("Latitude"), row.get("Longitude"))
    warnings.extend(coord_warns)
    for w in coord_warns:
        issue(
            "Latitude/Longitude",
            f"{row.get('Latitude')},{row.get('Longitude')}",
            w,
            "HIGH" if w == "INVALID_COORDINATES" else "MEDIUM",
            "N",
            loc or "",
            "WARN",
        )

    pt_raw = s(row.get("Property_Type")).lower()
    if pt_raw in PROPERTY_TYPE_MAP:
        ptype = PROPERTY_TYPE_MAP[pt_raw]
    elif pt_raw == "building":
        ptype = None
        reviews.append("UNKNOWN_PROPERTY_TYPE")
        warnings.append("UNKNOWN_PROPERTY_TYPE")
        issue("Property_Type", row.get("Property_Type"), "UNKNOWN_PROPERTY_TYPE", "HIGH", "N", "", "REVIEW")
    else:
        ptype = None
        reviews.append("UNKNOWN_PROPERTY_TYPE")
        warnings.append("UNKNOWN_PROPERTY_TYPE")
        issue("Property_Type", row.get("Property_Type") or "(blank)", "UNKNOWN_PROPERTY_TYPE", "HIGH", "N", "", "REVIEW")

    created_at = parse_datetime_iso(row.get("date_created"))
    updated_at = parse_datetime_iso(row.get("date_updated"))
    if not created_at:
        warnings.append("MISSING_SURVEY_DATE")
        issue("date_created", row.get("date_created"), "INVALID_DATE", "HIGH", "N", "", "WARN")

    owner = s(row.get("Owner_Name"))
    if owner and not name_ok(owner):
        issue("Owner_Name", owner, "INVALID_NAME", "MEDIUM", "N", "", "WARN")

    plot_val, plot_unit, plot_warn = parse_area_native(row.get("Plot_Area") or "")
    if plot_warn:
        warnings.append(plot_warn)
        issue("Plot_Area", row.get("Plot_Area"), plot_warn, "MEDIUM", "N", "", "WARN")

    gba, gba_warn = parse_sqft(row.get("Total_Area") or "")
    if gba_warn:
        warnings.append(gba_warn)
        issue(
            "Total_Area",
            row.get("Total_Area"),
            gba_warn,
            "LOW" if gba_warn == "AREA_APPROX" else "MEDIUM",
            "N",
            str(gba) if gba else "",
            "WARN",
        )

    amt, amt_warn = parse_asking_amount(row.get("Amount") or "")
    if amt_warn:
        reviews.append("UNPARSEABLE_ASKING_AMOUNT")
        warnings.append(amt_warn)
        issue("Amount", row.get("Amount"), amt_warn, "HIGH", "N", "", "REVIEW")

    nf = parse_int_field(row.get("No_of_Floors") or "")
    units_json, unit_warns = build_units(row)
    warnings.extend(unit_warns)

    market_json, _, competitors = build_market_details(row)
    brand_candidates = []
    unknown_brands = []
    for name in competitors:
        matched = match_business(name)
        if matched:
            brand_candidates.append(matched)
        else:
            unknown_brands.append(name)

    address = s(row.get("Address"))
    if not address:
        warnings.append("MISSING_ADDRESS")
        issue("Address", row.get("Address"), "MISSING_ADDRESS", "MEDIUM", "N", "", "WARN")

    availability = s(row.get("status"))
    possession = s(row.get("Possession"))
    if possession and len(possession) > 255:
        warnings.append("POSSESSION_EXCEEDS_VARCHAR255")
        issue("Possession", possession, "POSSESSION_TOO_LONG", "MEDIUM", "N", "", "WARN")

    an = s(row.get("Available_Not_Rented")).lower()
    if an == "yes":
        pstatus = "VACANT"
    elif an == "no":
        pstatus = "RENTED"
    else:
        pstatus = ""
    poss_l = possession.lower()
    if pstatus == "RENTED" and re.search(r"\bavailable\b", poss_l) and not re.search(r"not available|unavailable", poss_l):
        reviews.append("PROPERTY_STATUS_CONFLICT")
        warnings.append("PROPERTY_STATUS_CONFLICT")
        issue("Available_Not_Rented/Possession", f"{row.get('Available_Not_Rented')} | {possession}", "PROPERTY_STATUS_CONFLICT", "HIGH", "N", pstatus, "REVIEW")
    if pstatus == "VACANT" and re.search(r"\b(leased|rented|occupied)\b", poss_l):
        reviews.append("PROPERTY_STATUS_CONFLICT")
        warnings.append("PROPERTY_STATUS_CONFLICT")
        issue("Available_Not_Rented/Possession", f"{row.get('Available_Not_Rented')} | {possession}", "PROPERTY_STATUS_CONFLICT", "HIGH", "N", pstatus, "REVIEW")

    available_for = s(row.get("Available_For"))
    if available_for.lower() in {"select", "."}:
        available_for = ""

    bd = s(row.get("Business_Developer"))
    bd_email, bd_status = user_match(bd)
    if bd_status == "UNMATCHED":
        reviews.append("UNMATCHED_USER")
        warnings.append("UNMATCHED_USER")
        issue("Business_Developer", bd, "UNMATCHED_USER", "MEDIUM", "N", "", "REVIEW")

    for fld, dest, limit in (
        ("Advance", "advance", 255),
        ("Security", "security", 255),
        ("Grace_Period", "gracePeriod", 255),
        ("Increment", "increment", 255),
        ("Agreement_Period", "agreementPeriod", 255),
        ("Plot_Size", "plotSize", 100),
        ("Covered_Size", "coveredSize", 100),
        ("Available_For", "availableFor", 50),
    ):
        val = s(row.get(fld))
        if val and len(val) > limit:
            warnings.append(f"{dest.upper()}_EXCEEDS_VARCHAR{limit}")
            issue(fld, val, f"EXCEEDS_VARCHAR{limit}", "MEDIUM", "N", "", "WARN")

    notes = []
    if area_source:
        notes.append(f"areaSource={area_source}")
    if brand_candidates:
        notes.append("seedBusinessCandidates=" + ",".join(sorted(set(brand_candidates))))
    if unknown_brands:
        notes.append(f"unknownBrandCount={len(unknown_brands)}")
    if warnings:
        notes.append("warnings=" + ",".join(dict.fromkeys(warnings)))
    if reviews:
        notes.append("review=" + ",".join(dict.fromkeys(reviews)))

    status = "READY"
    if reviews:
        status = "REVIEW"
    elif warnings:
        status = "READY_WITH_WARNINGS"

    survey = {
        "legacySurveyId": lid,
        "migrationStatus": status,
        "migrationWarnings": ";".join(dict.fromkeys(warnings)),
        "migrationNotes": " | ".join(notes),
        "userMatchStatus": bd_status,
        "businessDeveloperEmail": bd_email or bd,
        "matchedUserId": "",
        "serialNumber": "",
        "propertyName": "",
        "address": address,
        "locationCoordinates": loc or "",
        "city": city or "",
        "area": area,
        "zipCode": "",
        "yearBuilt": "",
        "yearRenovated": "",
        "buildingClass": "",
        "propertyType": ptype or "",
        "ownerName": owner,
        "ownershipEntity": "",
        "propertyManager": "",
        "acquisitionDate": "",
        "plotArea": float(plot_val) if plot_val is not None else "",
        "areaUnit": plot_unit or "",
        "plotSize": s(row.get("Plot_Size")),
        "landShape": "",
        "topography": "",
        "zoning": "",
        "grossBuildingArea": float(gba) if gba is not None else "",
        "rentableArea": "",
        "usableArea": "",
        "coveredSize": s(row.get("Covered_Size")),
        "numberOfFloors": nf if nf is not None else "",
        "ceilingHeight": "",
        "constructionType": "",
        "exteriorMaterial": "",
        "roofType": "",
        "foundationType": "",
        "parkingSpaces": "",
        "loadingDocks": "",
        "elevatorCount": "",
        "hvacDetails": "",
        "fireProtection": "",
        "availabilityStatus": availability,
        "propertyStatus": pstatus,
        "availableFor": available_for,
        "askingAmount": float(amt) if amt is not None else "",
        "possession": possession,
        "survey": "",
        "advance": s(row.get("Advance")),
        "security": s(row.get("Security")),
        "gracePeriod": s(row.get("Grace_Period")),
        "increment": s(row.get("Increment")),
        "agreementPeriod": s(row.get("Agreement_Period")),
        "surveyDate": created_at,
        "financialIncome": "",
        "financialExpenses": "",
        "financialMetrics": "",
        "occupancySurvey": "",
        "leaseSurvey": "",
        "utilityDetails": "",
        "units": units_json or "",
        "siteDetails": "",
        "buildingCondition": "",
        "environmentalDetails": "",
        "legalDetails": "",
        "marketDetails": market_json or "",
        "attachments": "",
        "createdById": "",
        "updatedById": "",
        "createdAt": created_at,
        "updatedAt": updated_at,
        "freshSurveyCount": 0,
        "neighboringBusinessIds": "[]",
        "submittedAt": "",
    }
    return survey, issues, decisions


def extract_attachments(row: dict, head: dict[str, dict]) -> list[dict]:
    out = []
    lid = int(row["id"])
    for source in ("Pictures", "Videos", "Backup_Documents"):
        urls = parse_urls(row.get(source) or "")
        for idx, url in enumerate(urls):
            name, ext = filename_of(url)
            atype = classify_attachment_type(ext, source)
            exists, status, notes = attachment_status(url, head, atype)
            info = head.get(url) or {}
            size = ""
            mime = ""
            if info.get("content_length"):
                try:
                    size = int(info["content_length"])
                except ValueError:
                    size = ""
            if info.get("content_type"):
                mime = info["content_type"].split(";")[0]
            tag = "BACKUP" if source == "Backup_Documents" else ""
            out.append(
                {
                    "legacySurveyId": lid,
                    "legacyAttachmentSource": source,
                    "legacyAttachmentIndex": idx,
                    "originalUrl": url,
                    "originalPath": urlparse(url).path,
                    "originalFilename": name,
                    "extension": ext,
                    "mimeType": mime,
                    "fileSizeBytes": size,
                    "checksum": "",
                    "attachmentType": atype,
                    "attachmentTag": tag,
                    "sourceExists": exists,
                    "migrationStatus": status,
                    "targetObjectKey": "",
                    "newAttachmentRecordId": "",
                    "notes": notes,
                }
            )
    return out


def find_duplicates(rows: list[dict]) -> list[dict]:
    groups: dict[tuple, list[int]] = defaultdict(list)
    for r in rows:
        lat, lng = s(r.get("Latitude")), s(r.get("Longitude"))
        owner = re.sub(r"[^a-z0-9]", "", s(r.get("Owner_Name")).lower())
        ta = re.sub(r"[^a-z0-9]", "", s(r.get("Total_Area")).lower())
        amt = s(r.get("Amount"))
        if lat and lng and owner and ta and amt:
            groups[(lat, lng, owner, ta, amt)].append(int(r["id"]))
    out = []
    gid = 1
    for ids in groups.values():
        if len(ids) < 2:
            continue
        for sid in sorted(ids):
            out.append(
                {
                    "duplicateGroupId": gid,
                    "legacySurveyId": sid,
                    "duplicateReason": "same lat+lng+owner+Total_Area+Amount",
                    "confidence": "HIGH",
                    "keepForMigration": "TRUE",
                    "notes": "Re-survey is not automatically a duplicate. Review before merging.",
                }
            )
        gid += 1
    return out


def field_meanings() -> dict[str, str]:
    return {
        "id": "Gravity Forms entry ID",
        "form_id": "Gravity Forms form ID",
        "post_id": "WP post id (unused)",
        "date_created": "Entry created timestamp (0000-00-00 = missing)",
        "date_updated": "Entry updated (unused in this extract)",
        "is_starred": "GF starred flag",
        "is_read": "GF read flag",
        "ip": "Submitter IP",
        "source_url": "Form page URL",
        "user_agent": "Browser UA",
        "currency": "GF payment currency",
        "payment_status": "GF payment (unused)",
        "payment_date": "GF payment (unused)",
        "payment_amount": "GF payment (unused)",
        "payment_method": "GF payment (unused)",
        "transaction_id": "GF payment (unused)",
        "is_fulfilled": "GF payment (unused)",
        "created_by": "WordPress user id of the staff member who entered the form",
        "transaction_type": "GF payment (unused)",
        "status": "GF entry status (always active)",
        "Owner_Name": "Property owner / company name",
        "Contact_Person": "Role word or person name of the property contact",
        "Contact_Number": "Property contact phone(s)",
        "Business_Developer": "Internal Shapers BD email — NOT a property contact",
        "Address": "Street address",
        "City": "City",
        "Areas_In_Lahore": "Lahore locality dropdown",
        "Market_Area_114": "Lahore/ISB/RWP market free text",
        "Areas_In_Islamabad": "Islamabad locality dropdown",
        "Areas_In_Rawalpindi": "Rawalpindi locality dropdown",
        "Areas_In_Karachi": "Free-text area for non LHR/ISB/RWP cities (label is misleading)",
        "Market_Area_117": "Karachi market dropdown",
        "Longitude": "Longitude",
        "Latitude": "Latitude",
        "Map_Search_Address": "Google Maps formatted address",
        "Show_Map": "Map widget mode",
        "Map_Type": "Roadmap / Satellite",
        "Property_Type": "building | plot | office | house",
        "Plot_Size": "Plot dimensions text",
        "Plot_Area": "Plot area in native unit",
        "Add_Floor_Wise_Details": "Yes / No / Details Box",
        "Covered_Size": "Covered dimensions or size",
        "Select_Floors": "JSON array of floor labels offered",
        "Area_Details_Box": "HTML area narrative",
        "Total_Area": "Total area offered",
        "Pictures": "JSON array of image URLs",
        "Videos": "JSON array of video URLs",
        "Available_For": "Rent or Sale listing intent",
        "Rent_Sale_Price_Available": "Whether an asking amount was entered",
        "No_Rent_Price_Text": "HTML price note",
        "Amount": "Asking rent or sale price PKR",
        "Deal": "Price basis (Negotiable / Asking / Net)",
        "Advance": "Advance / prepaid rent text",
        "Security": "Security deposit text",
        "Possession": "Availability / possession phrasing",
        "Grace_Period": "Rent-free period",
        "Increment": "Rent escalation",
        "Agreement_Period": "Offered lease term",
        "Note": "HTML general notes",
        "Earlier_Brands": "Free-text nearby / earlier occupiers",
        "Select_Brands": "JSON brand checklist",
        "Other_Brands": "Free-text other brands",
        "Backup_Documents": "JSON array of extra file URLs (often images)",
        "Available_Not_Rented": "Yes/No checkbox",
        "Is_Approved": "GF approval code",
        "No_of_Floors": "Number of floors",
        "Floor_1_Type": "Floor slot 1 type",
        "Floor_1_Area": "Floor slot 1 area",
        "Floor_1_Size": "Floor slot 1 dimensions",
        "Floor_2_Type": "Floor slot 2 type",
        "Floor_2_Area": "Floor slot 2 area",
        "Floor_2_Size": "Floor slot 2 dimensions",
        "Floor_3_Type": "Floor slot 3 type",
        "Floor_3_Area": "Floor slot 3 area",
        "Floor_3_Size": "Floor slot 3 dimensions",
        "Floor_4_Type": "Floor slot 4 type",
        "Floor_4_Area": "Floor slot 4 area",
        "Floor_4_Size": "Floor slot 4 dimensions",
        "Floor_5_Type": "Floor slot 5 type",
        "Floor_5_Area": "Floor slot 5 area",
        "Floor_5_Size": "Floor slot 5 dimensions",
        "Floor_6_Type": "Floor slot 6 type",
        "Floor_6_Area": "Floor slot 6 area",
        "Floor_6_Size": "Floor slot 6 dimensions",
        "Floor_7_Type": "Floor slot 7 type",
        "Floor_7_Area": "Floor slot 7 area",
        "Floor_7_Size": "Floor slot 7 dimensions",
        "Floor_8_Type": "Floor slot 8 type",
        "Floor_8_Area": "Floor slot 8 area",
        "Floor_8_Size": "Floor slot 8 dimensions",
        "Floor_9_Type": "Floor slot 9 type",
        "Floor_9_Area": "Floor slot 9 area",
        "Floor_10_Type": "Floor slot 10 type",
        "Floor_10_Area": "Floor slot 10 area",
        "Floor_11_Type": "Floor slot 11 type",
        "Floor_11_Area": "Floor slot 11 area",
    }


def mapping_for(col: str) -> dict:
    meanings = field_meanings()
    # entity, destField, classification, transformation, confidence, decision
    table = {
        "id": ("manifest", "legacySurveyId", "MIGRATION_METADATA_ONLY", "Integer identity only; never Survey.id", "High", "APPROVED / FINAL"),
        "form_id": ("", "", "RAW_ONLY_SYSTEM_METADATA", "Gravity Forms meta", "High", "APPROVED / FINAL"),
        "post_id": ("", "", "RAW_ONLY_SYSTEM_METADATA", "Always empty", "High", "APPROVED / FINAL"),
        "date_created": ("Survey", "surveyDate + createdAt", "IMPORT_TO_SURVEY", "Valid dates only; 0000-00-00 → null", "High", "APPROVED / FINAL"),
        "date_updated": ("Survey", "updatedAt", "IMPORT_TO_SURVEY", "Valid dates only; currently empty in source", "High", "APPROVED / FINAL"),
        "is_starred": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "is_read": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "ip": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "source_url": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "user_agent": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "currency": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "payment_status": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "payment_date": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "payment_amount": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "payment_method": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "transaction_id": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "is_fulfilled": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "created_by": ("", "", "RAW_ONLY_SYSTEM_METADATA", "WP user id; not Survey.createdById", "High", "APPROVED / FINAL"),
        "transaction_type": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "status": ("Survey", "availabilityStatus", "IMPORT_TO_SURVEY", "Preserve source string (active)", "High", "APPROVED / FINAL"),
        "Owner_Name": ("Survey + SurveyContact", "ownerName / name", "IMPORT_TO_SURVEY", "Also SurveyContact.name when valid", "High", "APPROVED / FINAL"),
        "Contact_Person": ("SurveyContact", "role", "IMPORT_TO_SURVEY_CONTACT", "Pure roles only", "Medium", "APPROVED / FINAL"),
        "Contact_Number": ("SurveyContact", "phone", "IMPORT_TO_SURVEY_CONTACT", "One safe PK number", "High", "APPROVED / FINAL"),
        "Business_Developer": ("User", "createdById via email", "IMPORT_TO_USER_RELATION", "Never SurveyContact.email; matchedUserId null until CRE User lookup", "High", "APPROVED / FINAL"),
        "Address": ("Survey", "address", "IMPORT_TO_SURVEY", "Trim", "High", "APPROVED / FINAL"),
        "City": ("Survey", "city", "IMPORT_TO_SURVEY", "Spelling map; placeholders → null", "High", "APPROVED / FINAL"),
        "Areas_In_Lahore": ("Survey", "area", "IMPORT_TO_SURVEY", "When city=Lahore", "High", "APPROVED / FINAL"),
        "Market_Area_114": ("Survey", "area", "IMPORT_TO_SURVEY", "Fallback for LHR/ISB/RWP", "High", "APPROVED / FINAL"),
        "Areas_In_Islamabad": ("Survey", "area", "IMPORT_TO_SURVEY", "When city=Islamabad", "High", "APPROVED / FINAL"),
        "Areas_In_Rawalpindi": ("Survey", "area", "IMPORT_TO_SURVEY", "When city=Rawalpindi", "High", "APPROVED / FINAL"),
        "Areas_In_Karachi": ("Survey", "area", "IMPORT_TO_SURVEY", "Karachi + other-city free text", "High", "APPROVED / FINAL"),
        "Market_Area_117": ("Survey", "area", "IMPORT_TO_SURVEY", "Karachi market fallback", "High", "APPROVED / FINAL"),
        "Longitude": ("Survey", "locationCoordinates.longitude", "IMPORT_TO_SURVEY", "JSON pair with latitude", "High", "APPROVED / FINAL"),
        "Latitude": ("Survey", "locationCoordinates.latitude", "IMPORT_TO_SURVEY", "JSON pair with longitude", "High", "APPROVED / FINAL"),
        "Map_Search_Address": ("Survey", "area", "IMPORT_TO_SURVEY", "Last-resort area fallback only", "Medium", "APPROVED / FINAL"),
        "Show_Map": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "Map_Type": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "Property_Type": ("Survey", "propertyType", "IMPORT_TO_SURVEY", "plot/office/house mapped; building left null + REVIEW", "High", "APPROVED / FINAL"),
        "Plot_Size": ("Survey", "plotSize", "IMPORT_TO_SURVEY", "Preserve dimension string", "High", "APPROVED / FINAL"),
        "Plot_Area": ("Survey", "plotArea + areaUnit", "IMPORT_TO_SURVEY", "Parse only deterministic native units", "High", "APPROVED / FINAL"),
        "Add_Floor_Wise_Details": ("", "", "RAW_ONLY_REDUNDANT", "UI switch; floors themselves are imported", "High", "APPROVED / FINAL"),
        "Covered_Size": ("Survey", "coveredSize", "IMPORT_TO_SURVEY", "Preserve dimension string", "High", "APPROVED / FINAL"),
        "Select_Floors": ("Survey", "units[].floor", "IMPORT_TO_SURVEY_JSON", "Labels when no floor slots", "High", "APPROVED / FINAL"),
        "Area_Details_Box": ("", "", "UNMAPPED_BUSINESS_DATA", "HTML area narrative; no approved Survey field", "Medium", "PENDING_REVIEW"),
        "Total_Area": ("Survey", "grossBuildingArea", "IMPORT_TO_SURVEY", "Parse clean sq ft only", "High", "APPROVED / FINAL"),
        "Pictures": ("Survey", "attachments via 04", "IMPORT_TO_ATTACHMENTS", "S3 later", "High", "APPROVED / FINAL"),
        "Videos": ("Survey", "attachments via 04", "IMPORT_TO_ATTACHMENTS", "S3 later", "High", "APPROVED / FINAL"),
        "Available_For": ("Survey", "availableFor", "IMPORT_TO_SURVEY", "Rent / Sale", "High", "APPROVED / FINAL"),
        "Rent_Sale_Price_Available": ("", "", "RAW_ONLY_REDUNDANT", "Equals Amount filled", "High", "APPROVED / FINAL"),
        "No_Rent_Price_Text": ("", "", "UNMAPPED_BUSINESS_DATA", "No approved destination", "High", "PENDING_REVIEW"),
        "Amount": ("Survey", "askingAmount", "IMPORT_TO_SURVEY", "Numeric Decimal when parseable", "High", "APPROVED / FINAL"),
        "Deal": ("", "", "UNMAPPED_BUSINESS_DATA", "No approved destination", "High", "PENDING_REVIEW"),
        "Advance": ("Survey", "advance", "IMPORT_TO_SURVEY", "Preserve original string", "High", "APPROVED / FINAL"),
        "Security": ("Survey", "security", "IMPORT_TO_SURVEY", "Preserve original string", "High", "APPROVED / FINAL"),
        "Possession": ("Survey", "possession", "IMPORT_TO_SURVEY", "Preserve original string", "High", "APPROVED / FINAL"),
        "Grace_Period": ("Survey", "gracePeriod", "IMPORT_TO_SURVEY", "Preserve original string", "High", "APPROVED / FINAL"),
        "Increment": ("Survey", "increment", "IMPORT_TO_SURVEY", "Preserve original string", "High", "APPROVED / FINAL"),
        "Agreement_Period": ("Survey", "agreementPeriod", "IMPORT_TO_SURVEY", "Preserve original string", "High", "APPROVED / FINAL"),
        "Note": ("", "", "UNMAPPED_BUSINESS_DATA", "No approved destination", "High", "PENDING_REVIEW"),
        "Earlier_Brands": ("Survey", "marketDetails.nearbyCompetitors", "IMPORT_TO_SURVEY_JSON", "Also candidate for neighboringBusinessIds", "Medium", "APPROVED / FINAL"),
        "Select_Brands": ("Survey", "marketDetails.nearbyCompetitors", "IMPORT_TO_SURVEY_JSON", "Also candidate for neighboringBusinessIds", "Medium", "APPROVED / FINAL"),
        "Other_Brands": ("Survey", "marketDetails.nearbyCompetitors", "IMPORT_TO_SURVEY_JSON", "Also candidate for neighboringBusinessIds", "Medium", "APPROVED / FINAL"),
        "Backup_Documents": ("Survey", "attachments via 04", "IMPORT_TO_ATTACHMENTS", "OTHER type HOLD until CRE supports documents", "High", "APPROVED / FINAL"),
        "Available_Not_Rented": ("Survey", "propertyStatus", "IMPORT_TO_SURVEY", "Yes→VACANT No→RENTED; conflicts REVIEW", "High", "APPROVED / FINAL"),
        "Is_Approved": ("", "", "RAW_ONLY_SYSTEM_METADATA", "", "High", "APPROVED / FINAL"),
        "No_of_Floors": ("Survey", "numberOfFloors", "IMPORT_TO_SURVEY", "Integral only", "High", "APPROVED / FINAL"),
    }
    for i in FLOOR_SLOTS:
        table[f"Floor_{i}_Type"] = ("Survey", "units[]", "IMPORT_TO_SURVEY_JSON", "units.number / units.floor", "High", "APPROVED / FINAL")
        table[f"Floor_{i}_Area"] = ("Survey", "units[]", "IMPORT_TO_SURVEY_JSON", "units.squareFeet when parseable", "High", "APPROVED / FINAL")
        if i <= 8:
            table[f"Floor_{i}_Size"] = ("Survey", "units[].dimensions", "IMPORT_TO_SURVEY_JSON", "Preserved in units JSON", "High", "APPROVED / FINAL")

    m = table.get(col)
    if not m:
        return {
            "legacyField": col,
            "legacyMeaning": meanings.get(col, ""),
            "destinationEntity": "",
            "destinationField": "",
            "mappingType": "UNKNOWN",
            "transformation": "",
            "confidence": "",
            "migrationDecision": "PENDING_REVIEW",
            "notes": "",
        }
    ent, dest, mtype, trans, conf, dec = m
    return {
        "legacyField": col,
        "legacyMeaning": meanings.get(col, ""),
        "destinationEntity": ent,
        "destinationField": dest,
        "mappingType": mtype,
        "transformation": trans,
        "confidence": conf,
        "migrationDecision": dec,
        "notes": "",
    }


def unmapped_fields() -> list[dict]:
    return [
        {
            "legacyField": "No_Rent_Price_Text",
            "meaning": "HTML price note",
            "reasonNotMapped": "No approved Survey field",
            "possibleTarget": "notes or priceNote (not approved)",
            "schemaChangeRequired": "Y",
        },
        {
            "legacyField": "Deal",
            "meaning": "Price basis (Negotiable / Asking / Net)",
            "reasonNotMapped": "No approved Survey field",
            "possibleTarget": "priceBasis (not approved)",
            "schemaChangeRequired": "Y",
        },
        {
            "legacyField": "Note",
            "meaning": "HTML general notes",
            "reasonNotMapped": "No approved Survey field",
            "possibleTarget": "notes Text (not approved)",
            "schemaChangeRequired": "Y",
        },
        {
            "legacyField": "Area_Details_Box",
            "meaning": "HTML area narrative (110 rows)",
            "reasonNotMapped": "Discovered additional business text; not in the three named unresolved fields but has no approved destination",
            "possibleTarget": "notes (not approved)",
            "schemaChangeRequired": "Y",
        },
    ]
