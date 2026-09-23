#!/usr/bin/env python3
"""Generate the canonical legacy Survey migration workbook.

Usage (from this folder or the properties-viewer repo root):

    python migration/legacy-survey-migration/generate-workbook.py

Reads:
    migration/legacy-survey-migration/source/merged_properties_final.csv
    migration/legacy-survey-migration/source/attachment_head_scan.csv  (optional)

Writes:
    migration/legacy-survey-migration/legacy-survey-migration-7398.xlsx

Does not touch PostgreSQL, S3, or Prisma.
"""
from __future__ import annotations

import csv
import json
import sys
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).resolve().parent
SOURCE = HERE / "source" / "merged_properties_final.csv"
HEAD = HERE / "source" / "attachment_head_scan.csv"
OUTPUT = HERE / "legacy-survey-migration-7398.xlsx"

sys.path.insert(0, str(Path(__file__).resolve().parent))
from final_schema_map import write_sheet as write_final_schema_mapping  # noqa: E402
from normalize import (  # noqa: E402
    EXPECTED_ROW_COUNT,
    classify_contact,
    extract_attachments,
    find_duplicates,
    load_head_scan,
    mapping_for,
    normalize_city,
    normalize_survey,
    pick_area,
    unmapped_fields,
)

FONT = Font(name="Arial", size=10)
FONT_B = Font(name="Arial", size=10, bold=True)
FONT_H = Font(name="Arial", size=14, bold=True)
FONT_BLUE = Font(name="Arial", size=10, color="0000FF")
HEADER_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
YELLOW = PatternFill("solid", fgColor="FFFF00")
LEGEND = PatternFill("solid", fgColor="FFF2CC")
WRAP = Alignment(wrap_text=True, vertical="top")
TOP = Alignment(vertical="top")

SHEETS = [
    "00_Final_Schema_Mapping",
    "01_Legacy_Raw_Data",
    "02_Survey_Normalized",
    "03_Survey_Contacts",
    "04_Attachments",
    "05_Field_Mapping",
    "06_Enum_Mappings",
    "07_JSON_Mappings",
    "08_Unmapped_Fields",
    "09_Data_Issues",
    "10_Duplicates",
    "11_Decision_Register",
    "12_Migration_Summary",
]


def header_row(ws, headers, row=1):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row, c, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = f"A{row + 1}"
    ws.auto_filter.ref = f"A{row}:{get_column_letter(len(headers))}{row}"
    ws.row_dimensions[row].height = 28


def widths(ws, sizes):
    for i, w in enumerate(sizes, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_rows(ws, headers, rows, start=1, wrap_cols=None):
    header_row(ws, headers, start)
    wrap_cols = set(wrap_cols or [])
    for r_i, row in enumerate(rows, start + 1):
        values = row if isinstance(row, (list, tuple)) else [row.get(h, "") for h in headers]
        for c, val in enumerate(values, 1):
            cell = ws.cell(r_i, c, val if val != "" else None)
            cell.font = FONT
            cell.alignment = WRAP if c in wrap_cols else TOP
    if rows:
        ws.auto_filter.ref = f"A{start}:{get_column_letter(len(headers))}{start + len(rows)}"
    return start + len(rows)


def load_source(path: Path) -> list[dict]:
    if not path.exists():
        raise SystemExit(f"Source not found: {path}")
    with path.open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))
    if len(rows) != EXPECTED_ROW_COUNT:
        raise SystemExit(f"Expected {EXPECTED_ROW_COUNT} rows, found {len(rows)} in {path}")
    ids = [r.get("id") for r in rows]
    if len(ids) != len(set(ids)):
        raise SystemExit("Source contains duplicate ids")
    return rows


def examples(rows, col, n=3):
    seen = []
    for r in rows:
        v = (r.get(col) or "").strip()
        if v and v not in seen:
            seen.append(v[:80])
        if len(seen) >= n:
            break
    return " · ".join(seen)


def coverage(rows, col):
    return sum(1 for r in rows if (r.get(col) or "").strip())


def schema_decisions():
    approved = [
        ("AREA", "Areas_In_* / Market_Area_* / Map_Search_Address", "One Survey.area string; city-specific source then Map_Search_Address fallback"),
        ("COORDINATES", "Latitude+Longitude", "JSON pair in locationCoordinates; no separate lat/lng Survey columns"),
        ("STATUS", "status", "→ availabilityStatus"),
        ("PROPERTY_STATUS", "Available_Not_Rented", "Yes→VACANT, No→RENTED; conflicts REVIEW"),
        ("BD", "Business_Developer", "User.email lookup → createdById; unmatched REVIEW; never SurveyContact.email"),
        ("PLOT_SIZE", "Plot_Size", "→ plotSize raw string"),
        ("COVERED_SIZE", "Covered_Size", "→ coveredSize raw string"),
        ("AVAILABLE_FOR", "Available_For", "→ availableFor"),
        ("ASKING_AMOUNT", "Amount", "→ askingAmount Decimal when parseable"),
        ("ADVANCE", "Advance", "→ advance"),
        ("SECURITY", "Security", "→ security"),
        ("POSSESSION", "Possession", "→ possession"),
        ("GRACE", "Grace_Period", "→ gracePeriod"),
        ("INCREMENT", "Increment", "→ increment"),
        ("AGREEMENT", "Agreement_Period", "→ agreementPeriod"),
        ("DATE_UPDATED", "date_updated", "→ updatedAt"),
        ("DATE_CREATED", "date_created", "→ surveyDate + createdAt"),
        ("FLOORS", "Floor_1..11 + Select_Floors", "→ units JSON including dimensions"),
        ("BRANDS", "Earlier_Brands / Select_Brands / Other_Brands", "→ marketDetails.nearbyCompetitors; neighboringBusinessIds empty until CRE Business ids exist"),
        ("ATTACHMENTS", "Pictures / Videos / Backup_Documents", "04 manifest; upload later"),
        ("SURVEY_FIELD", "survey", "null — no legacy source found"),
        ("PROPERTY_TYPE", "Property_Type", "plot/office/house mapped; building left null + REVIEW"),
    ]
    rows = []
    for did, field, decision in approved:
        rows.append(
            {
                "decisionId": did,
                "category": "MAPPING",
                "legacySurveyId": "",
                "legacyField": field,
                "rawValue": "",
                "problem": "",
                "recommendedOptions": decision,
                "currentDecision": decision,
                "status": "APPROVED / FINAL",
                "notes": "",
            }
        )
    pending = [
        ("UNMAPPED-NO-RENT-PRICE-TEXT", "No_Rent_Price_Text", "HTML price note", "notes / priceNote / keep raw"),
        ("UNMAPPED-DEAL", "Deal", "Negotiable / Asking / Net", "priceBasis / keep raw"),
        ("UNMAPPED-NOTE", "Note", "HTML general notes", "notes Text / keep raw"),
        ("UNMAPPED-AREA-DETAILS-BOX", "Area_Details_Box", "HTML area narrative (110 rows)", "notes / keep raw"),
        ("CRE-LOCATION-COORDINATES-TYPE", "locationCoordinates", "CRE field is still Decimal(9,6)", "Change CRE column to Json before import"),
        ("CRE-TWELVE-FIELDS", "12 approved Survey fields", "Not yet in CRE Prisma", "Add on CRE before import"),
        ("CRE-UNITS-DIMENSIONS", "units.dimensions", "SurveyUnitDto has no dimensions key today", "Extend JSON DTO on CRE or extra keys will be stripped"),
        ("CRE-SURVEY-DATE-NULLABLE", "date_created", "1,491 zero dates; CRE surveyDate is NOT NULL", "Make nullable or HOLD those rows"),
        ("CRE-USER-LOOKUP", "Business_Developer", "No production User table in this workspace", "Importer must resolve createdById"),
        ("CRE-ATTACHMENT-OTHER", "Backup_Documents", "AttachmentRecord is IMAGE|VIDEO only", "Extend type or HOLD OTHER files"),
    ]
    for did, field, problem, options in pending:
        rows.append(
            {
                "decisionId": did,
                "category": "PENDING",
                "legacySurveyId": "",
                "legacyField": field,
                "rawValue": "",
                "problem": problem,
                "recommendedOptions": options,
                "currentDecision": "",
                "status": "PENDING_REVIEW",
                "notes": "",
            }
        )
    return rows


def build():
    print(f"Reading {SOURCE}")
    rows = load_source(SOURCE)
    keys = list(rows[0].keys())
    print(f"Source rows={len(rows)} columns={len(keys)}")
    head = load_head_scan(HEAD)
    print(f"HEAD scan urls={len(head)}")

    surveys = []
    contacts = []
    attachments = []
    issues = []
    decisions = schema_decisions()
    for r in rows:
        survey, iss, dec = normalize_survey(r)
        surveys.append(survey)
        issues.extend(iss)
        decisions.extend(dec)
        contacts.append(classify_contact(r))
        attachments.extend(extract_attachments(r, head))
    duplicates = find_duplicates(rows)

    map_rows = []
    for col in keys:
        meta = mapping_for(col)
        meta["legacyCoverage"] = f"{coverage(rows, col)}/{len(rows)}"
        meta["exampleValues"] = examples(rows, col)
        map_rows.append(meta)

    # enums
    pt_counts = Counter((r.get("Property_Type") or "").strip() or "(blank)" for r in rows)
    city_counts = Counter((r.get("City") or "").strip() for r in rows)
    af_counts = Counter((r.get("Available_For") or "").strip() or "(blank)" for r in rows)
    an_counts = Counter((r.get("Available_Not_Rented") or "").strip() or "(blank)" for r in rows)
    deal_counts = Counter((r.get("Deal") or "").strip() or "(blank)" for r in rows)
    role_counts = Counter((r.get("Contact_Person") or "").strip() for r in rows)
    enum_rows = []
    pt_norm = {"plot": "LAND_PLOT", "office": "OFFICE", "house": "RESIDENTIAL", "building": "", "(blank)": ""}
    for val, n in pt_counts.most_common():
        enum_rows.append(["PropertyType", val, n, pt_norm.get(val.lower() if val != "(blank)" else "(blank)", ""), "HIGH" if val in ("plot", "office", "house") else "LOW", "Y" if val in ("building", "(blank)") else "N", "building/blank left null"])
    for val, n in af_counts.most_common():
        enum_rows.append(["availableFor", val, n, val if val not in ("(blank)",) else "", "HIGH", "N", "Available_For → availableFor (APPROVED)"])
    for val, n in an_counts.most_common():
        mapped_an = {"Yes": "VACANT", "No": "RENTED", "(blank)": ""}.get(val, "")
        enum_rows.append(["Available_Not_Rented", val, n, mapped_an, "HIGH", "N" if mapped_an else "Y", "→ propertyStatus (APPROVED)"])
    for val, n in deal_counts.most_common():
        enum_rows.append(["Deal", val, n, "", "N/A", "Y", "No priceBasis field"])
    for val in ("Select", ".", "jhalum", "Jehlum", "Muree", "Burawala", "huripur", "Wah", "rawat", "hattar"):
        n = city_counts.get(val, 0)
        if n:
            mapped = { "Select": "", ".": "", "jhalum": "Jhelum", "Jehlum": "Jhelum", "Muree": "Murree", "Burawala": "Burewala", "huripur": "Haripur", "Wah": "Wah", "rawat": "Rawat", "hattar": "Hattar" }[val]
            enum_rows.append(["City", val, n, mapped, "MEDIUM", "Y" if val in ("Select", ".", "Wah") else "N", "Wah kept distinct from Wah Cantt."])
    for val, n in role_counts.most_common(25):
        from normalize import PURE_ROLES
        enum_rows.append(["ContactRole", val, n, PURE_ROLES.get(val.lower(), ""), "HIGH" if val.lower() in PURE_ROLES else "LOW", "N" if val.lower() in PURE_ROLES else "Y", ""])
    enum_rows.append(["BuildingClass", "(no source)", len(rows), "", "N/A", "N", "Leave null"])
    enum_rows.append(["availabilityStatus", "active (from status)", len(rows), "active", "HIGH", "N", "status → availabilityStatus (APPROVED)"])
    enum_rows.append(["PropertyStatus", "Available_Not_Rented=Yes", sum(1 for r in rows if (r.get("Available_Not_Rented") or "").strip().lower() == "yes"), "VACANT", "HIGH", "N", "APPROVED"])
    enum_rows.append(["PropertyStatus", "Available_Not_Rented=No", sum(1 for r in rows if (r.get("Available_Not_Rented") or "").strip().lower() == "no"), "RENTED", "HIGH", "N", "APPROVED"])
    enum_rows.append(["areaUnit", "Kanal (from Plot_Area)", sum(1 for r in rows if "kanal" in (r.get("Plot_Area") or "").lower()), "Kanal", "HIGH", "N", "Native unit kept"])
    enum_rows.append(["areaUnit", "Marla (from Plot_Area)", sum(1 for r in rows if "marla" in (r.get("Plot_Area") or "").lower()), "Marla", "HIGH", "N", "Native unit kept"])

    # JSON mappings
    units_n = sum(1 for s in surveys if s["units"])
    market_n = sum(1 for s in surveys if s["marketDetails"])
    json_rows = [
        ["Select_Floors + Floor_1..11 Type/Area/Size", "units", "floor, squareFeet, dimensions", units_n, "Property Viewer units JSON includes dimensions when Floor_*_Size is present, e.g. {\"floor\":\"Ground Floor\",\"squareFeet\":2700,\"dimensions\":\"032ft x 085ft\"}. CRE SurveyUnitDto does not yet whitelist dimensions — extend DTO later. Do not drop dimensions from this workbook.", "N"],
        ["Select_Brands + Other_Brands + Earlier_Brands", "marketDetails", "nearbyCompetitors[]", market_n, "Strings only; neighboringBusinessIds stays []", "Y"],
        ["Amount", "financialIncome", "rentalIncome", 0, "NOT mapped — asking price is not income", "Y"],
        ["Advance/Security/Agreement", "leaseSurvey", "baseRent/leaseTerm", 0, "NOT mapped — offer ≠ existing lease", "Y"],
        ["(none)", "financialExpenses", "", 0, "No source", "Y"],
        ["(none)", "financialMetrics", "", 0, "Server-computed — leave empty", "Y"],
        ["(none)", "occupancySurvey", "", 0, "Server-computed from units — leave empty for importer", "Y"],
        ["(none)", "utilityDetails", "", 0, "No boolean utility source", "Y"],
        ["(none)", "siteDetails", "", 0, "No source", "Y"],
        ["(none)", "buildingCondition", "", 0, "No condition ratings", "Y"],
        ["(none)", "environmentalDetails", "", 0, "No source", "Y"],
        ["(none)", "legalDetails", "", 0, "No source", "Y"],
        ["Pictures/Videos", "attachments", "AttachmentRecord[]", 0, "Blank on 02; real rows on 04; write after S3", "Y"],
    ]

    unmapped = []
    for u in unmapped_fields():
        field = u["legacyField"]
        # coverage for simple field names
        pop = coverage(rows, field) if field in keys else ""
        unmapped.append(
            {
                "legacyField": field,
                "meaning": u["meaning"],
                "populatedRows": pop if pop != "" else "see notes",
                "examples": examples(rows, field) if field in keys else "",
                "reasonNotMapped": u["reasonNotMapped"],
                "possibleTarget": u["possibleTarget"],
                "schemaChangeRequired": u["schemaChangeRequired"],
                "decisionStatus": "PENDING_REVIEW",
                "notes": "Original values remain on 01_Legacy_Raw_Data",
            }
        )

    area_sources = Counter()
    for r in rows:
        city, _ = normalize_city(r.get("City"))
        _, src = pick_area(r, city)
        area_sources[src or "(none)"] += 1
    class_counts = Counter(m["mappingType"] for m in map_rows)

    wb = Workbook()

    ws0 = wb.active
    ws0.title = "00_Final_Schema_Mapping"
    schema00 = write_final_schema_mapping(ws0, rows, surveys)

    # 01
    ws = wb.create_sheet("01_Legacy_Raw_Data")
    ws.sheet_properties.tabColor = "1F4E79"
    header_row(ws, keys)
    for i, r in enumerate(rows, 2):
        for c, k in enumerate(keys, 1):
            val = r.get(k, "")
            cell = ws.cell(i, c, val if val != "" else None)
            cell.font = FONT
            cell.alignment = TOP
    for c, k in enumerate(keys, 1):
        ws.column_dimensions[get_column_letter(c)].width = min(max(len(k) + 2, 12), 28)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(keys))}{1 + len(rows)}"
    ws.freeze_panes = "B2"
    ws["A1"].comment = Comment(
        "Immutable raw extract. Source: merged_properties_final.csv (7,398 × 96). "
        "Do not edit. properties.json is a 53-column viewer subset and is not the source.",
        "Migration generator",
    )

    # 02
    ws = wb.create_sheet("02_Survey_Normalized")
    ws.sheet_properties.tabColor = "2E7D32"
    h2 = list(surveys[0].keys())
    write_rows(ws, h2, surveys, wrap_cols={3, 4, 9, 47})
    widths(ws, [16, 22, 36, 40] + [16] * 20)
    ws.freeze_panes = "B2"

    # 03
    ws = wb.create_sheet("03_Survey_Contacts")
    h3 = [
        "legacySurveyId", "legacyContactSource", "legacyContactIndex",
        "name", "role", "phone", "email",
        "legacyName", "legacyRole", "legacyPhone",
        "validationStatus", "migrationStatus", "migrationNotes",
    ]
    write_rows(ws, h3, contacts, wrap_cols={13})
    widths(ws, [16, 28, 12, 22, 18, 18, 14, 22, 22, 22, 22, 18, 40])

    # 04
    ws = wb.create_sheet("04_Attachments")
    h4 = list(attachments[0].keys()) if attachments else [
        "legacySurveyId", "legacyAttachmentSource", "legacyAttachmentIndex", "originalUrl"
    ]
    write_rows(ws, h4, attachments, wrap_cols={4, 16})
    widths(ws, [14, 22, 12, 55, 36, 28, 10, 16, 12, 12, 12, 14, 12, 14, 16, 16, 36])

    # 05
    ws = wb.create_sheet("05_Field_Mapping")
    h5 = [
        "legacyField", "legacyMeaning", "legacyCoverage", "exampleValues",
        "destinationEntity", "destinationField", "mappingType", "transformation",
        "confidence", "migrationDecision", "notes",
    ]
    write_rows(ws, h5, map_rows, wrap_cols={2, 4, 8})
    widths(ws, [28, 40, 14, 40, 28, 32, 20, 40, 12, 18, 20])

    # 06
    ws = wb.create_sheet("06_Enum_Mappings")
    write_rows(
        ws,
        ["mappingGroup", "legacyValue", "recordCount", "normalizedValue", "confidence", "reviewRequired", "notes"],
        enum_rows,
        wrap_cols={7},
    )
    widths(ws, [22, 28, 12, 18, 12, 14, 48])

    # 07
    ws = wb.create_sheet("07_JSON_Mappings")
    write_rows(
        ws,
        ["legacyField", "surveyJsonField", "nestedKey", "populatedNormalizedRows", "rule", "whitelistSafe"],
        json_rows,
        wrap_cols={5},
    )
    widths(ws, [40, 22, 28, 16, 52, 14])

    # 08
    ws = wb.create_sheet("08_Unmapped_Fields")
    h8 = [
        "legacyField", "meaning", "populatedRows", "examples", "reasonNotMapped",
        "possibleTarget", "schemaChangeRequired", "decisionStatus", "notes",
    ]
    write_rows(ws, h8, unmapped, wrap_cols={2, 4, 5})
    widths(ws, [36, 28, 14, 36, 48, 28, 14, 16, 36])

    # 09
    ws = wb.create_sheet("09_Data_Issues")
    h9 = [
        "legacySurveyId", "field", "rawValue", "issueType", "severity",
        "automaticFixPossible", "normalizedValue", "migrationAction", "notes",
    ]
    write_rows(ws, h9, issues, wrap_cols={3, 9})
    widths(ws, [16, 22, 40, 28, 10, 14, 20, 14, 28])

    # 10
    ws = wb.create_sheet("10_Duplicates")
    h10 = ["duplicateGroupId", "legacySurveyId", "duplicateReason", "confidence", "keepForMigration", "notes"]
    write_rows(ws, h10, duplicates, wrap_cols={6})
    widths(ws, [16, 16, 36, 12, 18, 52])

    # 11
    ws = wb.create_sheet("11_Decision_Register")
    ws.sheet_properties.tabColor = "FFFF00"
    ws["A1"] = "Edit only currentDecision (col H) and status (col I). Yellow + blue = inputs. Do not invent business data."
    ws["A1"].font = FONT_B
    ws["A1"].fill = LEGEND
    ws["A2"] = "Example: set I5 to APPROVED after accepting the recommendation in H5. Default status is PENDING_REVIEW."
    ws["A2"].font = FONT
    ws["A2"].fill = LEGEND
    h11 = [
        "decisionId", "category", "legacySurveyId", "legacyField", "rawValue",
        "problem", "recommendedOptions", "currentDecision", "status", "notes",
    ]
    write_rows(ws, h11, decisions, start=4, wrap_cols={5, 6, 7, 8, 10})
    widths(ws, [28, 22, 14, 28, 40, 40, 40, 40, 16, 28])
    last = 4 + len(decisions)
    for r in range(5, last + 1):
        for col in (8, 9):
            ws.cell(r, col).font = FONT_BLUE
            ws.cell(r, col).fill = YELLOW
    dv = DataValidation(type="list", formula1='"PENDING_REVIEW,APPROVED,REJECTED"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"I5:I{last}")

    # 12 summary — static values computed from the same in-memory rows (no Excel formulas)
    status_c = Counter(s["migrationStatus"] for s in surveys)
    contact_action = Counter(c["migrationAction"] for c in contacts)
    contact_status = Counter(c["validationStatus"] for c in contacts)
    att_src = Counter(a["legacyAttachmentSource"] for a in attachments)
    att_type = Counter(a["attachmentType"] for a in attachments)
    att_st = Counter(a["migrationStatus"] for a in attachments)
    att_ex = Counter(a["sourceExists"] for a in attachments)
    coord_ok = sum(1 for s in surveys if s["locationCoordinates"])
    coord_missing = sum(1 for s in surveys if "MISSING_COORDINATES" in s["migrationWarnings"])
    coord_invalid = sum(1 for s in surveys if "INVALID_COORDINATES" in s["migrationWarnings"])
    pt_mapped = sum(1 for s in surveys if s["propertyType"])
    pt_unresolved = len(surveys) - pt_mapped
    avail_c = Counter(s["availabilityStatus"] or "(blank)" for s in surveys)
    pstat_c = Counter(s["propertyStatus"] or "(blank)" for s in surveys)
    user_c = Counter(s["userMatchStatus"] for s in surveys)
    conflict_n = sum(1 for s in surveys if "PROPERTY_STATUS_CONFLICT" in s["migrationWarnings"])
    floor_any = 0
    floor_4 = 0
    floor_9 = 0
    highest = 0
    for r in rows:
        used = []
        for i in range(1, 12):
            if (r.get(f"Floor_{i}_Type") or "").strip() or (r.get(f"Floor_{i}_Area") or "").strip() or (
                r.get(f"Floor_{i}_Size") or ""
            ).strip():
                used.append(i)
        if used:
            floor_any += 1
            highest = max(highest, max(used))
            if any(i >= 4 for i in used):
                floor_4 += 1
            if any(i >= 9 for i in used):
                floor_9 += 1

    ws = wb.create_sheet("12_Migration_Summary")
    ws["A1"] = "Legacy Survey Migration Workbook"
    ws["A1"].font = FONT_H
    lines = [
        (3, "Source file", str(SOURCE)),
        (4, "Verified source rows", len(rows)),
        (5, "Source columns", len(keys)),
        (6, "Normalized Survey rows", len(surveys)),
        (7, "READY", status_c.get("READY", 0)),
        (8, "READY_WITH_WARNINGS", status_c.get("READY_WITH_WARNINGS", 0)),
        (9, "REVIEW", status_c.get("REVIEW", 0)),
        (10, "HOLD", status_c.get("HOLD", 0)),
        (12, "Contact candidates", len(contacts)),
        (13, "Contacts MIGRATE", contact_action.get("MIGRATE", 0)),
        (14, "Contacts MIGRATE_PHONE_NULL", contact_action.get("MIGRATE_PHONE_NULL", 0)),
        (15, "Contacts HOLD", contact_action.get("HOLD", 0)),
        (16, "VALID", contact_status.get("VALID", 0)),
        (17, "VALID_NO_PHONE", contact_status.get("VALID_NO_PHONE", 0)),
        (18, "REVIEW_MULTI_PHONE", contact_status.get("REVIEW_MULTI_PHONE", 0)),
        (19, "INVALID_NAME", contact_status.get("INVALID_NAME", 0)),
        (20, "REVIEW_CP_SEMANTICS", contact_status.get("REVIEW_CP_SEMANTICS", 0)),
        (22, "Attachment references", len(attachments)),
        (23, "Pictures", att_src.get("Pictures", 0)),
        (24, "Videos", att_src.get("Videos", 0)),
        (25, "Backup_Documents", att_src.get("Backup_Documents", 0)),
        (26, "IMAGE type", att_type.get("IMAGE", 0)),
        (27, "VIDEO type", att_type.get("VIDEO", 0)),
        (28, "OTHER type", att_type.get("OTHER", 0)),
        (29, "Attachments READY", att_st.get("READY", 0)),
        (30, "Attachments REVIEW", att_st.get("REVIEW", 0)),
        (31, "Attachments HOLD", att_st.get("HOLD", 0)),
        (32, "sourceExists=Y", att_ex.get("Y", 0)),
        (33, "sourceExists=N", att_ex.get("N", 0)),
        (34, "sourceExists=UNKNOWN", att_ex.get("UNKNOWN", 0)),
        (36, "Valid locationCoordinates JSON", coord_ok),
        (37, "Missing coordinates", coord_missing),
        (38, "Invalid coordinates", coord_invalid),
        (40, "PropertyType mapped", pt_mapped),
        (41, "PropertyType unresolved", pt_unresolved),
        (42, "availabilityStatus=active", avail_c.get("active", 0)),
        (43, "propertyStatus=VACANT", pstat_c.get("VACANT", 0)),
        (44, "propertyStatus=RENTED", pstat_c.get("RENTED", 0)),
        (45, "propertyStatus blank", pstat_c.get("(blank)", 0)),
        (46, "propertyStatus conflicts", conflict_n),
        (48, "area from Areas_In_Lahore", area_sources.get("Areas_In_Lahore", 0)),
        (49, "area from Areas_In_Islamabad", area_sources.get("Areas_In_Islamabad", 0)),
        (50, "area from Areas_In_Rawalpindi", area_sources.get("Areas_In_Rawalpindi", 0)),
        (51, "area from Areas_In_Karachi", area_sources.get("Areas_In_Karachi", 0)),
        (52, "area from Market_Area_114", area_sources.get("Market_Area_114", 0)),
        (53, "area from Market_Area_117", area_sources.get("Market_Area_117", 0)),
        (54, "area from Map_Search_Address", area_sources.get("Map_Search_Address", 0)),
        (55, "area none", area_sources.get("(none)", 0)),
        (57, "BD MATCH_SEED_CANDIDATE", user_c.get("MATCH_SEED_CANDIDATE", 0)),
        (58, "BD UNMATCHED", user_c.get("UNMATCHED", 0)),
        (60, "Surveys with floor slots", floor_any),
        (61, "Highest floor slot used", highest),
        (62, "Surveys with Floor 4+", floor_4),
        (63, "Surveys with Floor 9+", floor_9),
        (65, "Duplicate member rows", len(duplicates)),
        (66, "Duplicate groups", len({d["duplicateGroupId"] for d in duplicates})),
        (67, "Data-issue rows", len(issues)),
        (68, "Decision rows", len(decisions)),
        (69, "Unmapped field groups", len(unmapped)),
        (70, "Legacy columns on 05", len(map_rows)),
        (72, "IMPORT_TO_SURVEY", class_counts.get("IMPORT_TO_SURVEY", 0)),
        (73, "IMPORT_TO_SURVEY_JSON", class_counts.get("IMPORT_TO_SURVEY_JSON", 0)),
        (74, "IMPORT_TO_SURVEY_CONTACT", class_counts.get("IMPORT_TO_SURVEY_CONTACT", 0)),
        (75, "IMPORT_TO_ATTACHMENTS", class_counts.get("IMPORT_TO_ATTACHMENTS", 0)),
        (76, "IMPORT_TO_RELATION", class_counts.get("IMPORT_TO_RELATION", 0)),
        (77, "IMPORT_TO_USER_RELATION", class_counts.get("IMPORT_TO_USER_RELATION", 0)),
        (78, "MIGRATION_METADATA_ONLY", class_counts.get("MIGRATION_METADATA_ONLY", 0)),
        (79, "RAW_ONLY_SYSTEM_METADATA", class_counts.get("RAW_ONLY_SYSTEM_METADATA", 0)),
        (80, "RAW_ONLY_REDUNDANT", class_counts.get("RAW_ONLY_REDUNDANT", 0)),
        (81, "UNMAPPED_BUSINESS_DATA", class_counts.get("UNMAPPED_BUSINESS_DATA", 0)),
    ]
    ws["A2"] = (
        "Counts are computed by the generator from the 7,398 source rows. "
        "CRE ERP was not modified. Nothing was imported or uploaded."
    )
    ws["A2"].font = FONT
    ws["A2"].alignment = WRAP
    ws.row_dimensions[2].height = 32
    for r, label, val in lines:
        ws.cell(r, 1, label).font = FONT_B
        ws.cell(r, 2, val).font = FONT
    ws["A83"] = "CRE changes still required before import (not done here)"
    ws["A83"].font = FONT_B
    ws["A84"] = "1) Add the 12 approved Survey columns to Prisma."
    ws["A85"] = "2) Change locationCoordinates from Decimal(9,6) to Json {latitude,longitude}."
    ws["A86"] = "3) Extend SurveyUnitDto with optional dimensions."
    ws["A87"] = "4) Make surveyDate nullable or HOLD 1,491 zero-date rows."
    ws["A88"] = "5) Importer User.email lookup for createdById; Business.id lookup for neighboringBusinessIds."
    ws["A89"] = "6) Decide IMAGE|VIDEO-only vs OTHER Backup_Documents."
    for r in range(84, 90):
        ws.cell(r, 1).font = FONT
    widths(ws, [78, 70])

    for s in wb.worksheets:
        s.page_setup.orientation = "landscape"
        s.page_setup.fitToPage = True
        s.page_setup.fitToWidth = 1
        s.page_setup.fitToHeight = 0
        s.sheet_view.showGridLines = True
        s.oddHeader.left.text = '&"Arial"Legacy Survey Migration 7398'
        s.oddFooter.right.text = '&"Arial"Page &P of &N'

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    out = OUTPUT
    print(f"Writing {out}")
    try:
        wb.save(out)
    except PermissionError:
        out = HERE / "legacy-survey-migration-7398-approved.xlsx"
        print(f"Primary workbook is locked; writing {out}")
        wb.save(out)
    print("Wrote", out, "size", out.stat().st_size)

    summary = {
        "source": str(SOURCE),
        "rows": len(rows),
        "columns": len(keys),
        "surveyStatus": dict(status_c),
        "contacts": dict(contact_action),
        "contactStatus": dict(contact_status),
        "attachments": len(attachments),
        "attachmentSource": dict(att_src),
        "attachmentType": dict(att_type),
        "attachmentStatus": dict(att_st),
        "coordinates": {"valid": coord_ok, "missing": coord_missing, "invalid": coord_invalid},
        "propertyType": {"mapped": pt_mapped, "unresolved": pt_unresolved},
        "duplicates": {"members": len(duplicates), "groups": len({d["duplicateGroupId"] for d in duplicates})},
        "areaSources": dict(area_sources),
        "availabilityStatus": dict(avail_c),
        "propertyStatus": dict(pstat_c),
        "propertyStatusConflicts": conflict_n,
        "userMatch": dict(user_c),
        "floors": {"withSlots": floor_any, "highest": highest, "floor4plus": floor_4, "floor9plus": floor_9},
        "classifications": dict(class_counts),
        "issues": len(issues),
        "decisions": len(decisions),
        "sheet00": schema00,
        "output": str(out),
    }
    (OUTPUT.parent / "generation-summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(json.dumps(summary, indent=2))
    return summary


if __name__ == "__main__":
    build()
