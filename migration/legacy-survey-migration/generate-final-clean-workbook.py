#!/usr/bin/env python3
"""Generate the clean legacy Survey migration workbook (data only, no analysis).

Usage (from the repo root):

    python migration/legacy-survey-migration/generate-final-clean-workbook.py

Reads:
    migration/legacy-survey-migration/source/merged_properties_final.csv
    migration/legacy-survey-migration/source/attachment_head_scan.csv

Writes:
    migration/legacy-survey-migration/legacy-survey-final-clean-7398.xlsx
    migration/legacy-survey-migration/final-clean-summary.json

Sheets: 01_Surveys, 02_Survey_Contacts, 03_Attachments, 04_Exceptions.

Transformation rules come from normalize.py (the same code behind the audit
workbook legacy-survey-migration-7398.xlsx). This script only adds:
  * full timestamps for createdAt (normalize.py keeps the date only)
  * measurements with an "(approx.)" qualifier or thousands separators, parsed by
    stripping the qualifier and handing the remainder to normalize.py's parsers
  * units JSON as {floor, squareFeet?, dimensions?} with no duplicated "number"
  * Select_Floors cells merged as '[...] | [...]'
  * comma-separated brand lists split into separate names when unambiguous
  * createdById = the record's own Business_Developer email (importer maps it to User.id)
  * one SurveyContact row per survey with the record's raw Owner_Name / Contact_Person /
    Contact_Number; validation findings are warnings, never exclusions
Everything that cannot enter the clean sheets goes to 04_Exceptions.
Every value comes from the same legacy id; nothing is borrowed across records.

Does not touch the audit workbook, PostgreSQL, S3 or the CRE project.
"""
from __future__ import annotations

import sys

sys.dont_write_bytecode = True  # keep the tracked __pycache__ untouched

import csv  # noqa: E402
import json  # noqa: E402
import re  # noqa: E402
from collections import Counter  # noqa: E402
from datetime import datetime  # noqa: E402
from decimal import Decimal  # noqa: E402
from pathlib import Path  # noqa: E402

from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import normalize as N  # noqa: E402

SOURCE = HERE / "source" / "merged_properties_final.csv"
HEAD = HERE / "source" / "attachment_head_scan.csv"
OUTPUT = HERE / "legacy-survey-final-clean-7398.xlsx"
SUMMARY = HERE / "final-clean-summary.json"

SURVEY_FIELDS = [
    "serialNumber", "propertyName", "address", "locationCoordinates", "city", "area",
    "zipCode", "yearBuilt", "yearRenovated", "buildingClass", "propertyType", "ownerName",
    "ownershipEntity", "propertyManager", "acquisitionDate", "plotArea", "areaUnit",
    "plotSize", "landShape", "topography", "zoning", "grossBuildingArea", "rentableArea",
    "usableArea", "coveredSize", "numberOfFloors", "ceilingHeight", "constructionType",
    "exteriorMaterial", "roofType", "foundationType", "parkingSpaces", "loadingDocks",
    "elevatorCount", "hvacDetails", "fireProtection", "availabilityStatus",
    "propertyStatus", "availableFor", "askingAmount", "possession", "survey", "advance",
    "security", "gracePeriod", "increment", "agreementPeriod", "surveyDate",
    "financialIncome", "financialExpenses", "financialMetrics", "occupancySurvey",
    "leaseSurvey", "utilityDetails", "units", "siteDetails", "buildingCondition",
    "environmentalDetails", "legalDetails", "marketDetails", "attachments", "createdById",
    "updatedById", "createdAt", "updatedAt", "freshSurveyCount", "neighboringBusinessIds",
    "submittedAt",
]
SURVEY_HEADERS = ["legacySurveyId"] + SURVEY_FIELDS
CONTACT_HEADERS = ["legacySurveyId", "name", "role", "phone", "email", "createdAt"]
ATTACHMENT_HEADERS = [
    "legacySurveyId", "attachmentType", "originalUrl", "originalFilename", "mimeType",
    "fileSizeBytes",
]
EXCEPTION_HEADERS = [
    "entity", "legacySurveyId", "field", "rawValue", "normalizedValue", "reason",
    "severity", "notes",
]
PENDING_FIELDS = ["No_Rent_Price_Text", "Deal", "Note", "Area_Details_Box"]
IMPORTABLE_CONTACT = {"MIGRATE", "MIGRATE_PHONE_NULL"}
CONTACT_SOURCE = {"name": "Owner_Name", "role": "Contact_Person", "phone": "Contact_Number"}
MIME_PREFIX = {"IMAGE": "image/", "VIDEO": "video/"}
AREA_COLUMNS = [
    "Areas_In_Lahore", "Market_Area_114", "Areas_In_Islamabad", "Areas_In_Rawalpindi",
    "Areas_In_Karachi", "Market_Area_117", "Map_Search_Address",
]
# "Nil"/"None" are kept: in lease terms they assert "none required".
PLACEHOLDER_TOKENS = {".", "-", "--", "n/a", "na", "null", "undefined", "select"}


# ---------------------------------------------------------------- helpers

def parse_timestamp(raw: str) -> datetime | None:
    v = N.s(raw)
    if not v or v.startswith("0000-00-00"):
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(v, fmt)
        except ValueError:
            continue
    return None


APPROX_RE = re.compile(r"\(?\s*\bapprox(?:imately|\.)?\s*[.,]?\s*\)?", re.I)
NUMBER_RE = re.compile(r"\d[\d,]*(?:\.\d+)?")
GROUPED_RE = re.compile(r"^(?:\d{1,3}(?:,\d{3})+|\d{1,2}(?:,\d{2})*,\d{3})(?:\.\d+)?$")
SQFT_RE = re.compile(r"\(?\s*(?:sq\s*[.\-]?\s*f(?:ee)?t\.?|sqft\.?|sft\.?)\s*\)?", re.I)


def _strip_measure(raw: str) -> tuple[str | None, bool]:
    """Remove an approx qualifier and valid thousands separators.

    Returns (cleaned text, approx flag) or (None, approx) when a number uses an
    ambiguous comma such as "3,36".
    """
    v = N.s(raw)
    approx = bool(re.search(r"approx", v, re.I))
    v = APPROX_RE.sub(" ", v)
    for num in NUMBER_RE.findall(v):
        if "," in num:
            if not GROUPED_RE.match(num):
                return None, approx
            v = v.replace(num, num.replace(",", ""), 1)
    return re.sub(r"\s+", " ", v).strip(" .,"), approx


def parse_sqft_clean(raw: str) -> tuple[Decimal | None, bool]:
    """(value, approx). normalize.parse_sqft first, then qualifier-stripped retry."""
    val, warn = N.parse_sqft(raw)
    if val is not None:
        return val, warn == "AREA_APPROX"
    cleaned, approx = _strip_measure(raw)
    if not cleaned:
        return None, approx
    cleaned = SQFT_RE.sub(" sqft", cleaned).strip()
    val, _ = N.parse_sqft(cleaned)
    return val, approx


def parse_plot_area_clean(raw: str) -> tuple[Decimal | None, str | None, bool]:
    """(value, unit, approx). normalize.parse_area_native first, then retry."""
    val, unit, _ = N.parse_area_native(raw)
    if val is not None:
        return val, unit, False
    cleaned, approx = _strip_measure(raw)
    if not cleaned:
        return None, None, approx
    val, unit, _ = N.parse_area_native(cleaned)
    return val, unit, approx


def number_out(d: Decimal | float | None):
    if d is None or d == "":
        return None
    d = Decimal(str(d))
    return int(d) if d == d.to_integral_value() else float(d)


def parse_json_multi(raw: str) -> tuple[list, bool]:
    """JSON list, or several JSON lists joined by ' | ' (merge artefact).

    Returns (items, fully_parsed).
    """
    text = N.s(raw)
    if not text:
        return [], True
    items: list = []
    ok = True
    for part in text.split(" | "):
        part = part.strip()
        if part in ("", "[]"):
            continue
        parsed = N.parse_json_list(part)
        if not parsed:
            ok = False
        items.extend(parsed)
    return items, ok


POSITIONAL_RE = re.compile(
    r"^(?:near(?:by)?|besides?|opposite|opp\b\.?|adjacent|adj\b\.?|next\s+to|in\s*front\s+of|"
    r"front\s+of|behind|facing|close\s+to|along)\b",
    re.I,
)


def split_brands(text: str) -> tuple[list[str], bool]:
    """Split a delimiter-separated brand list. (names, split_was_safe)."""
    t = re.sub(r"[,\s]*\betc\b\.?\s*$", "", N.s(text), flags=re.I).strip(" ,;.")
    if not t:
        return [], True
    if not re.search(r"[,;\n]", t):
        return [t], True
    pieces = [p.strip() for p in re.split(r"[,;\n]", t)]
    pieces = [p for p in pieces if p]
    for p in pieces:
        if (
            POSITIONAL_RE.search(p)
            or re.match(r"^[a-z]\b", p)          # "Salt,n Pepper", "Max,s Fast Food"
            or len(p) < 2
            or len(p.split()) > 6
            or p.count("(") != p.count(")")
        ):
            return [N.s(text)], False
    return pieces, True


def dumps(obj) -> str:
    return N.dumps(obj)


# ---------------------------------------------------------------- build

def build():
    with SOURCE.open(encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        rows = list(reader)
        columns = reader.fieldnames
    if len(rows) != N.EXPECTED_ROW_COUNT or len(columns) != 96:
        raise SystemExit(f"Unexpected source shape {len(rows)} x {len(columns)}")
    if len({r["id"] for r in rows}) != len(rows):
        raise SystemExit("Duplicate legacy ids in source")
    head = N.load_head_scan(HEAD)

    surveys, contacts, attachments, exceptions = [], [], [], []
    seen_urls: set[tuple[int, str]] = set()

    def exc(entity, lid, field, raw, norm, reason, severity, notes=""):
        exceptions.append({
            "entity": entity,
            "legacySurveyId": lid,
            "field": field,
            "rawValue": "" if raw is None else str(raw),
            "normalizedValue": "" if norm is None else str(norm),
            "reason": reason,
            "severity": severity,
            "notes": notes,
        })

    # normalize.py issue types that carry over one-to-one
    issue_map = {
        "PLACEHOLDER_CITY": ("city", "PLACEHOLDER_CITY", "MEDIUM", "Placeholder city left blank"),
        "CITY_SPELLING": ("city", "CITY_SPELLING_NORMALIZED", "LOW", "Spelling corrected"),
        "MISSING_COORDINATES": ("locationCoordinates", "MISSING_COORDINATES", "MEDIUM", ""),
        "INVALID_COORDINATES": ("locationCoordinates", "INVALID_COORDINATES", "HIGH", "Left blank"),
        "COORD_OUTSIDE_PAKISTAN": ("locationCoordinates", "COORDINATES_OUTSIDE_PAKISTAN", "MEDIUM", "Kept; verify"),
        "INVALID_NAME": ("ownerName", "INVALID_OWNER_NAME", "LOW", "Kept as ownerName per approved direct mapping; not used as SurveyContact.name"),
        "UNPARSEABLE_ASKING_AMOUNT": ("askingAmount", "UNPARSEABLE_ASKING_AMOUNT", "HIGH", "Left blank"),
        "POSSESSION_TOO_LONG": ("possession", "VALUE_EXCEEDS_COLUMN_LENGTH", "HIGH", "Kept in full; exceeds 255 chars"),
    }

    for row in rows:
        lid = int(row["id"])
        base, issues, _ = N.normalize_survey(row)
        rec = {f: base.get(f, "") for f in SURVEY_FIELDS}

        for it in issues:
            itype = it["issueType"]
            if itype in issue_map:
                field, reason, sev, note = issue_map[itype]
                exc("Survey", lid, field, it["rawValue"], it["normalizedValue"], reason, sev, note)
            elif itype.startswith("EXCEEDS_VARCHAR"):
                exc("Survey", lid, it["field"], it["rawValue"], "", "VALUE_EXCEEDS_COLUMN_LENGTH", "HIGH",
                    f"Kept in full; {itype}")
            elif itype == "PROPERTY_STATUS_CONFLICT":
                exc("Survey", lid, "propertyStatus", it["rawValue"], it["normalizedValue"],
                    "CONFLICT_WITH_POSSESSION", "HIGH",
                    "propertyStatus kept per approved Available_Not_Rented mapping")
            elif itype == "UNKNOWN_PROPERTY_TYPE":
                raw_pt = N.s(row.get("Property_Type"))
                if raw_pt:
                    exc("Survey", lid, "propertyType", raw_pt, "", "AMBIGUOUS_PROPERTY_TYPE", "MEDIUM",
                        "No safe PropertyType enum; left blank")
                else:
                    exc("Survey", lid, "propertyType", "", "", "MISSING_PROPERTY_TYPE", "MEDIUM",
                        "Property_Type blank in source")
            elif itype == "INVALID_DATE":
                exc("Survey", lid, "surveyDate/createdAt", it["rawValue"], "", "INVALID_LEGACY_DATE", "HIGH",
                    "surveyDate and createdAt left blank")
            # MISSING_ADDRESS: no source value, nothing lost.
            # UNMATCHED_USER: seed-list heuristic only; see createdById rows below.
            # AMBIGUOUS_AREA / AREA_APPROX: re-evaluated below with the clean parsers.

        # --- dates
        ts = parse_timestamp(row.get("date_created"))
        rec["surveyDate"] = ts.date() if ts else None
        rec["createdAt"] = ts
        rec["updatedAt"] = parse_timestamp(row.get("date_updated"))

        # --- plot area
        raw_pa = N.s(row.get("Plot_Area"))
        if raw_pa:
            val, unit, approx = parse_plot_area_clean(raw_pa)
            rec["plotArea"], rec["areaUnit"] = number_out(val), unit or ""
            if val is None:
                exc("Survey", lid, "plotArea", raw_pa, "", "UNPARSEABLE_PLOT_AREA", "MEDIUM",
                    "No single deterministic native unit; plotArea/areaUnit left blank")
            elif approx:
                exc("Survey", lid, "plotArea", raw_pa, f"{number_out(val)} {unit}", "APPROXIMATE_PLOT_AREA", "LOW",
                    "Stated value written; source marks it approximate")
        else:
            rec["plotArea"], rec["areaUnit"] = None, ""

        # --- gross building area
        raw_ta = N.s(row.get("Total_Area"))
        if raw_ta:
            val, approx = parse_sqft_clean(raw_ta)
            rec["grossBuildingArea"] = number_out(val)
            if val is None:
                exc("Survey", lid, "grossBuildingArea", raw_ta, "", "UNPARSEABLE_TOTAL_AREA", "MEDIUM",
                    "Not a single sq ft value; left blank")
            elif approx:
                exc("Survey", lid, "grossBuildingArea", raw_ta, number_out(val), "APPROXIMATE_TOTAL_AREA", "LOW",
                    "Stated sq ft written; source marks it approximate")
        else:
            rec["grossBuildingArea"] = None

        # --- units
        labels, labels_ok = parse_json_multi(row.get("Select_Floors"))
        labels = [N.s(x) for x in labels if N.s(x)]
        if not labels_ok:
            exc("Survey", lid, "units", row.get("Select_Floors"), dumps(labels), "UNPARSEABLE_SELECT_FLOORS",
                "MEDIUM", "Part of Select_Floors is not a JSON list")
        slots = []
        for i in N.FLOOR_SLOTS:
            ftype = N.s(row.get(f"Floor_{i}_Type"))
            farea = N.s(row.get(f"Floor_{i}_Area"))
            fsize = N.s(row.get(f"Floor_{i}_Size"))
            if ftype or farea or fsize:
                slots.append((i, ftype, farea, fsize))
        units = []
        if slots:
            unlabeled = [s for s in slots if not s[1]]
            positional = (
                unlabeled
                and len(unlabeled) == len(slots)
                and [s[0] for s in slots] == list(range(1, len(slots) + 1))
                and len(labels) == len(slots)
            )
            for pos, (i, ftype, farea, fsize) in enumerate(slots):
                unit = {}
                if ftype:
                    unit["floor"] = ftype
                elif positional:
                    unit["floor"] = labels[pos]
                    exc("Survey", lid, "units", f"Floor_{i} (no Floor_{i}_Type); Select_Floors={row.get('Select_Floors')}",
                        labels[pos], "UNIT_FLOOR_LABEL_FROM_SELECT_FLOORS", "LOW",
                        "Label taken positionally: slot count equals Select_Floors label count")
                else:
                    exc("Survey", lid, "units", f"Floor_{i}_Area={farea}; Floor_{i}_Size={fsize}", "",
                        "UNIT_FLOOR_LABEL_MISSING", "MEDIUM", "Unit written without a floor label")
                if farea:
                    sqft, approx = parse_sqft_clean(farea)
                    if sqft is not None:
                        unit["squareFeet"] = number_out(sqft)
                        if approx:
                            exc("Survey", lid, "units", f"Floor_{i}_Area={farea}", number_out(sqft),
                                "APPROXIMATE_FLOOR_AREA", "LOW", "Stated sq ft written")
                    else:
                        exc("Survey", lid, "units", f"Floor_{i}_Area={farea}", "", "UNPARSEABLE_FLOOR_AREA",
                            "MEDIUM", "squareFeet omitted for this unit")
                if fsize:
                    unit["dimensions"] = fsize
                units.append(unit)
        elif labels:
            units = [{"floor": lab} for lab in labels]
        rec["units"] = dumps(units) if units else ""

        # --- market details
        names = []
        brands, brands_ok = parse_json_multi(row.get("Select_Brands"))
        if not brands_ok:
            exc("Survey", lid, "marketDetails", row.get("Select_Brands"), "", "UNPARSEABLE_SELECT_BRANDS",
                "MEDIUM", "Part of Select_Brands is not a JSON list")
        names.extend(N.s(b) for b in brands if isinstance(b, str) and N.s(b))
        for col in ("Other_Brands", "Earlier_Brands"):
            text = N.s(row.get(col))
            if not text:
                continue
            parts, safe = split_brands(text)
            names.extend(parts)
            if not safe:
                exc("Survey", lid, "marketDetails", text, text, "MARKET_TEXT_NOT_SPLIT", "LOW",
                    f"{col} kept as one string (narrative or ambiguous delimiter)")
        seen, unique = set(), []
        for n in names:
            if n.lower() not in seen:
                seen.add(n.lower())
                unique.append(n)
        rec["marketDetails"] = dumps({"nearbyCompetitors": unique}) if unique else ""

        # --- relations / system defaults
        # createdById deliberately holds this record's Business_Developer email; the
        # CRE importer swaps it for User.id (User.email match). Never created_by.
        bd = N.s(row.get("Business_Developer"))
        if bd and "@" in bd:
            rec["createdById"] = bd
        else:
            rec["createdById"] = None
            exc("Survey", lid, "createdById", bd, "", "BUSINESS_DEVELOPER_MISSING", "MEDIUM",
                "No Business_Developer email on this record; createdById left blank")
        rec["attachments"] = ""
        rec["neighboringBusinessIds"] = "[]"
        rec["freshSurveyCount"] = 0

        if not rec["area"]:
            present = {c: N.s(row.get(c)) for c in AREA_COLUMNS if N.s(row.get(c))}
            if present:
                exc("Survey", lid, "area", dumps(present), "", "PLACEHOLDER_AREA", "LOW",
                    "Area columns hold only placeholders (Select / Area Not Listed / .)")

        for f in ("askingAmount", "numberOfFloors"):
            rec[f] = number_out(rec[f]) if rec[f] != "" else None

        # --- literal "missing" placeholders copied through by direct mappings
        for f in SURVEY_FIELDS:
            val = rec[f]
            if isinstance(val, str) and val.strip().lower() in PLACEHOLDER_TOKENS:
                exc("Survey", lid, f, val, "", "PLACEHOLDER_VALUE_BLANKED", "LOW",
                    "Placeholder meaning 'no value'; cell left blank")
                rec[f] = ""

        # --- pending business data
        for field in PENDING_FIELDS:
            raw = N.s(row.get(field))
            if raw:
                exc("Survey", lid, field, raw, "", "PENDING_BUSINESS_DECISION", "MEDIUM",
                    "No approved Survey destination")

        surveys.append({"legacySurveyId": lid, **rec})

        # --- contact: one row per survey, this record's own raw values.
        # Validation findings are warnings in 04_Exceptions, never exclusions.
        contact = {"legacySurveyId": lid, "email": "", "createdAt": ts}
        for dest, col in CONTACT_SOURCE.items():
            raw = N.s(row.get(col))
            if raw.lower() in PLACEHOLDER_TOKENS:
                exc("SurveyContact", lid, dest, raw, "", "PLACEHOLDER_VALUE_BLANKED", "LOW",
                    f"{col} placeholder meaning 'no value'; cell left blank")
                raw = ""
            contact[dest] = raw
        contacts.append(contact)

        c = N.classify_contact(row)
        if c["migrationStatus"] not in IMPORTABLE_CONTACT:
            exc("SurveyContact", lid, "contact",
                dumps({"Owner_Name": c["legacyName"], "Contact_Person": c["legacyRole"],
                       "Contact_Number": c["legacyPhone"]}),
                dumps({"name": c["name"], "role": c["role"], "phone": c["phone"]}),
                c["validationStatus"], "MEDIUM",
                f"Previously HOLD: {c['migrationNotes']}. Row kept in 02_Survey_Contacts with raw values")
        if not N.name_ok(c["legacyName"]):
            exc("SurveyContact", lid, "name", c["legacyName"], "", "CONTACT_NAME_INVALID", "LOW",
                "Title/placeholder rather than a name; raw value kept")
        tokens = N.phones(c["legacyPhone"])
        if len(tokens) >= 2:
            exc("SurveyContact", lid, "phone", c["legacyPhone"],
                dumps([N.normalize_phone(t) or t for t in tokens]), "CONTACT_MULTIPLE_PHONES", "LOW",
                "Full original string kept in phone; numbers not split")
        elif c["legacyPhone"] and not (tokens and N.normalize_phone(tokens[0])):
            exc("SurveyContact", lid, "phone", c["legacyPhone"], "", "CONTACT_PHONE_INCOMPLETE_OR_INVALID",
                "LOW", "Not a complete Pakistani number; raw value kept")

        # --- attachments
        for a in N.extract_attachments(row, head):
            mime = a["mimeType"]
            atype = a["attachmentType"]
            where = f"{a['legacyAttachmentSource']}[{a['legacyAttachmentIndex']}]"
            if a["migrationStatus"] == "READY":
                if mime and not mime.startswith(MIME_PREFIX.get(atype, "~")):
                    exc("Attachment", lid, where, a["originalUrl"], atype, "ATTACHMENT_MIME_MISMATCH", "MEDIUM",
                        f"Server returned {mime} for a {atype} reference")
                    continue
                if (lid, a["originalUrl"]) in seen_urls:
                    exc("Attachment", lid, where, a["originalUrl"], atype, "ATTACHMENT_DUPLICATE_REFERENCE", "LOW",
                        "Same URL already listed for this survey; imported once")
                    continue
                seen_urls.add((lid, a["originalUrl"]))
                attachments.append({
                    "legacySurveyId": lid,
                    "attachmentType": atype,
                    "originalUrl": a["originalUrl"],
                    "originalFilename": a["originalFilename"],
                    "mimeType": mime,
                    "fileSizeBytes": a["fileSizeBytes"] if a["fileSizeBytes"] != "" else None,
                })
                continue
            notes = a["notes"]
            if "Placeholder" in notes:
                reason, sev = "ATTACHMENT_PLACEHOLDER_PATH", "LOW"
            elif atype == "OTHER":
                reason, sev = "ATTACHMENT_TYPE_NOT_SUPPORTED", "MEDIUM"
            elif a["sourceExists"] == "UNKNOWN":
                reason, sev = "ATTACHMENT_SOURCE_NOT_VERIFIED", "MEDIUM"
            else:
                reason, sev = "ATTACHMENT_SOURCE_UNAVAILABLE", "MEDIUM"
            exc("Attachment", lid, where, a["originalUrl"], atype, reason, sev,
                f"{a['migrationStatus']}: {notes}".strip(": "))

    for d in N.find_duplicates(rows):
        exc("Survey", d["legacySurveyId"], "(record)", "", f"group {d['duplicateGroupId']}",
            "DUPLICATE_CANDIDATE", "MEDIUM", d["duplicateReason"] + "; both rows kept")

    status_values = Counter(N.s(r.get("status")) for r in rows)
    dataset = [
        ("availabilityStatus", dumps(dict(status_values)), "active", "LEGACY_STATUS_CONSTANT_ACTIVE", "MEDIUM",
         f"status → availabilityStatus kept as approved, but all {len(rows):,} legacy rows are 'active' "
         "(Gravity Forms entry status), so it carries no availability information"),
        ("createdById", "Business_Developer", "email", "CREATED_BY_ID_REQUIRES_USER_LOOKUP", "HIGH",
         "createdById holds each record's Business_Developer email, not an Int; importer must replace it "
         "with CRE User.id (match on User.email) before insert"),
        ("createdAt", "date_created", "", "TIMESTAMP_TIMEZONE_UNCONFIRMED", "MEDIUM",
         "Timestamps written as stored by Gravity Forms (normally UTC) with no offset; confirm before import"),
    ]
    for field, raw, norm, reason, sev, notes in dataset:
        exc("Dataset", None, field, raw, norm, reason, sev, notes)

    order = {"Dataset": 0, "Survey": 1, "SurveyContact": 2, "Attachment": 3}
    exceptions.sort(key=lambda e: (order[e["entity"]], e["legacySurveyId"] or 0, e["field"], e["reason"]))

    write_workbook(surveys, contacts, attachments, exceptions)

    summary = {
        "source": SOURCE.name,
        "sourceRows": len(rows),
        "sourceColumns": len(columns),
        "workbook": OUTPUT.name,
        "surveys": len(surveys),
        "contacts": len(contacts),
        "attachments": len(attachments),
        "attachmentTypes": dict(Counter(a["attachmentType"] for a in attachments)),
        "exceptions": len(exceptions),
        "exceptionsByEntity": dict(Counter(e["entity"] for e in exceptions)),
        "exceptionsByReason": dict(Counter(e["reason"] for e in exceptions).most_common()),
        "exceptionsByField": dict(Counter(e["field"] for e in exceptions if e["entity"] != "Attachment").most_common()),
    }
    SUMMARY.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(json.dumps(summary, indent=2))


# ---------------------------------------------------------------- write

HEADER_FONT = Font(bold=True)
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
TOP = Alignment(vertical="top")
WRAP = Alignment(vertical="top", wrap_text=True)

DATE_FMT = "yyyy-mm-dd"
DATETIME_FMT = "yyyy-mm-dd hh:mm:ss"


def write_sheet(wb, title, headers, records, widths=None, wrap=(), date_cols=(), datetime_cols=()):
    ws = wb.create_sheet(title)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    wrap, date_cols, datetime_cols = set(wrap), set(date_cols), set(datetime_cols)
    for r, rec in enumerate(records, 2):
        for c, h in enumerate(headers, 1):
            val = rec.get(h)
            if val == "":
                val = None
            cell = ws.cell(r, c)
            cell.value = val
            if isinstance(val, str):
                cell.data_type = "s"  # never let a leading "=" become a formula
            if h in date_cols:
                cell.number_format = DATE_FMT
            elif h in datetime_cols:
                cell.number_format = DATETIME_FMT
            cell.alignment = WRAP if h in wrap else TOP
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(len(records) + 1, 1)}"
    for c, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(c)].width = (widths or {}).get(h, max(12, min(len(h) + 4, 24)))
    return ws


def write_workbook(surveys, contacts, attachments, exceptions):
    wb = Workbook()
    wb.remove(wb.active)
    survey_json = {"locationCoordinates", "units", "marketDetails"}
    write_sheet(
        wb, "01_Surveys", SURVEY_HEADERS, surveys,
        widths={"address": 32, "locationCoordinates": 34, "area": 28, "ownerName": 26,
                "possession": 28, "units": 48, "marketDetails": 48, "createdById": 30},
        wrap=survey_json | {"address", "possession"},
        date_cols={"surveyDate"}, datetime_cols={"createdAt", "updatedAt"},
    )
    write_sheet(wb, "02_Survey_Contacts", CONTACT_HEADERS, contacts,
                widths={"name": 28, "role": 24, "phone": 28, "createdAt": 20},
                datetime_cols={"createdAt"})
    write_sheet(wb, "03_Attachments", ATTACHMENT_HEADERS, attachments,
                widths={"originalUrl": 70, "originalFilename": 36, "mimeType": 16})
    write_sheet(wb, "04_Exceptions", EXCEPTION_HEADERS, exceptions,
                widths={"entity": 14, "field": 24, "rawValue": 60, "normalizedValue": 30,
                        "reason": 34, "severity": 10, "notes": 50},
                wrap={"rawValue", "normalizedValue", "notes"})
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = 90
    print(f"Writing {OUTPUT}")
    wb.save(OUTPUT)


if __name__ == "__main__":
    if len(sys.argv) > 1:  # optional output path, e.g. when the workbook is open in Excel
        OUTPUT = Path(sys.argv[1]).resolve()
    build()
