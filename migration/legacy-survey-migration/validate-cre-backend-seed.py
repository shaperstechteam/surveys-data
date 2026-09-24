#!/usr/bin/env python3
"""Validate cre-backend-seed.xlsx against its spec and the final clean workbook.

Usage (from the repo root):

    python migration/legacy-survey-migration/validate-cre-backend-seed.py

Prints a JSON report and exits 1 on any error. The expected columns are written
out here (not imported from the generator) so the generator cannot validate itself.
"""
from __future__ import annotations

import sys

sys.dont_write_bytecode = True

import importlib.util  # noqa: E402
import json  # noqa: E402
import re  # noqa: E402
from collections import Counter, defaultdict  # noqa: E402
from datetime import datetime  # noqa: E402
from pathlib import Path  # noqa: E402

from openpyxl import load_workbook  # noqa: E402

HERE = Path(__file__).resolve().parent
SEED = HERE / "cre-backend-seed.xlsx"
CLEAN = HERE / "legacy-survey-final-clean-7398.xlsx"

SHEETS = ["01_Surveys", "02_Survey_Floors", "03_Survey_Contacts", "04_Attachments", "05_Seed_Exceptions"]
SURVEY_COLUMNS = """legacySurveyId
serialNumber propertyName address locationCoordinates city area zipCode yearBuilt yearRenovated
buildingClass propertyType
ownerName ownershipEntity propertyManager acquisitionDate
landArea landAreaUnit landFrontageFt landDepthFt landShape topography zoning
grossBuildingArea rentableArea usableArea numberOfFloors
constructionType exteriorMaterial roofType foundationType parkingSpaces loadingDocks elevatorCount
hvacDetails fireProtection
utilityPhase sanctionedLoadKva transformerAvailable transformerCapacityKva
availabilityStatus propertyStatus availableFor
advanceAmount securityDepositAmount possessionTerms
surveyDate
financialIncome financialExpenses financialMetrics occupancySurvey leaseSurvey utilityDetails units
siteDetails buildingCondition environmentalDetails legalDetails legalDocumentChecklist marketDetails
attachments
createdById updatedById createdAt updatedAt freshSurveyCount neighboringBusinessIds submittedAt""".split()
# NOTE: matches the CRE-oriented names requested for this workbook (floorWidthFt/floorDepthFt/
# floorAreaSqFt). If the actual Prisma SurveyFloor model still uses lengthFt/widthFt/areaSqFt,
# reconcile before import - this validator only checks internal consistency, not the live schema.
FLOOR_COLUMNS = ["legacySurveyId", "floorCode", "floorName", "sortOrder", "floorWidthFt", "floorDepthFt",
                 "floorAreaSqFt", "rentableAreaSqFt", "usableAreaSqFt", "ceilingHeightFt", "floorUse",
                 "notes", "attachments"]
CONTACT_COLUMNS = ["legacySurveyId", "name", "role", "phone", "email", "createdAt"]
ATTACHMENT_COLUMNS = ["legacySurveyId", "attachmentType", "originalUrl", "originalFilename", "mimeType",
                      "fileSizeBytes"]
EXCEPTION_COLUMNS = ["entity", "legacySurveyId", "field", "rawValue", "normalizedValue", "reason",
                     "severity", "notes"]
FORBIDDEN = ["askingAmount", "possession", "survey", "advance", "security", "gracePeriod", "increment",
             "agreementPeriod", "plotArea", "areaUnit", "plotSize", "coveredSize", "ceilingHeight"]
NOT_ALLOWED = {
    "id", "latitude", "longitude", "Latitude", "Longitude", "migrationStatus", "migrationWarnings",
    "migrationNotes", "userMatchStatus", "businessDeveloperEmail", "matchedUserId", "Business_Developer",
    "legacyContactSource", "legacyContactIndex", "legacyName", "legacyRole", "legacyPhone",
    "validationStatus", "sourceExists", "targetObjectKey", "checksum", "newAttachmentRecordId",
    # Gravity Forms / system columns
    "form_id", "post_id", "is_starred", "is_read", "ip", "source_url", "user_agent", "currency",
    "payment_status", "payment_date", "payment_amount", "payment_method", "transaction_id",
    "is_fulfilled", "created_by", "transaction_type", "Show_Map", "Map_Type", "Add_Floor_Wise_Details",
    "Rent_Sale_Price_Available", "Is_Approved", "status", "date_created", "date_updated",
}
JSON_COLUMNS = ["locationCoordinates", "financialIncome", "financialExpenses", "financialMetrics",
                "occupancySurvey", "leaseSurvey", "utilityDetails", "units", "siteDetails",
                "buildingCondition", "environmentalDetails", "legalDetails", "legalDocumentChecklist",
                "marketDetails", "attachments", "neighboringBusinessIds"]
EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
FLOOR_CODE_RE = re.compile(r"^(B[12]|LG|UG|G|M|RF|\d{1,2}F)$")


def read(wb, name):
    it = wb[name].iter_rows(values_only=True)
    headers = list(next(it))
    return headers, [dict(zip(headers, r)) for r in it if any(v is not None for v in r)]


def blank(v):
    return v is None or (isinstance(v, str) and not v.strip())


def main():
    errors = []

    def check(ok, msg):
        if not ok:
            errors.append(msg)

    seed = load_workbook(SEED)                      # full mode: formulas and merged cells visible
    check(seed.sheetnames == SHEETS, f"sheets {seed.sheetnames} != {SHEETS}")
    formulas = sum(1 for ws in seed.worksheets for row in ws.iter_rows() for c in row if c.data_type == "f")
    check(formulas == 0, f"{formulas} formula cells")
    check(sum(len(ws.merged_cells.ranges) for ws in seed.worksheets) == 0, "merged cells present")
    clean = load_workbook(CLEAN, read_only=True)

    spec = importlib.util.spec_from_file_location("seedgen", HERE / "generate-cre-backend-seed.py")
    gen = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(gen)

    data = {n: read(seed, n) for n in SHEETS}
    expected = {"01_Surveys": SURVEY_COLUMNS, "02_Survey_Floors": FLOOR_COLUMNS,
                "03_Survey_Contacts": CONTACT_COLUMNS, "04_Attachments": ATTACHMENT_COLUMNS,
                "05_Seed_Exceptions": EXCEPTION_COLUMNS}
    for n in SHEETS:
        h = data[n][0]
        check(h == expected[n], f"{n}: columns differ from spec")
        bad = [c for c in h if c in NOT_ALLOWED or (n not in ("02_Survey_Floors", "05_Seed_Exceptions")
                                                     and c in FORBIDDEN)]
        check(not bad, f"{n}: forbidden/helper/system columns {bad}")

    # ------------------------------------------------ surveys
    s_head, surveys = data["01_Surveys"]
    ids = [r["legacySurveyId"] for r in surveys]
    check(len(SURVEY_COLUMNS) == 68 and len(s_head) == 68, f"01_Surveys has {len(s_head)} columns, expected 68")
    check(len(surveys) == 7398, f"01_Surveys rows={len(surveys)}")
    check(len(set(ids)) == 7398, "legacySurveyId not unique")
    absent = [f for f in FORBIDDEN if f in s_head]
    check(not absent, f"removed fields present: {absent}")
    json_ok = Counter()
    for r in surveys:
        lid = r["legacySurveyId"]
        for c in JSON_COLUMNS:
            if not blank(r[c]):
                try:
                    json.loads(r[c])
                    json_ok[c] += 1
                except (TypeError, ValueError):
                    errors.append(f"{lid}: {c} is not valid JSON")
        check(r["neighboringBusinessIds"] == "[]", f"{lid}: neighboringBusinessIds != []")
        if not blank(r["createdById"]):
            check(isinstance(r["createdById"], str) and EMAIL.match(r["createdById"]),
                  f"{lid}: createdById {r['createdById']!r} is not an email")
        check(blank(r["createdAt"]) or isinstance(r["createdAt"], datetime), f"{lid}: createdAt not a timestamp")
        check(blank(r["surveyDate"]) or isinstance(r["surveyDate"], datetime), f"{lid}: surveyDate not a timestamp")
        for f in ("advanceAmount", "securityDepositAmount", "landFrontageFt", "landDepthFt"):
            check(blank(r[f]) or isinstance(r[f], (int, float)), f"{lid}: {f} {r[f]!r} is not numeric")
        check(blank(r["landFrontageFt"]) and blank(r["landDepthFt"]),
              f"{lid}: landFrontageFt/landDepthFt must stay null (frontage/depth order is unprovable)")
        if not blank(r["units"]):
            try:
                for item in json.loads(r["units"]):
                    check(isinstance(item, dict) and "dimensions" not in item and gen.UNIT_SIGNAL_KEYS & item.keys(),
                          f"{lid}: units item {item!r} has no genuine unit signal (looks like leftover floor data)")
            except (TypeError, ValueError):
                pass

    # every kept/derived value traces back to the final clean workbook, row by row on legacySurveyId
    _, c_surveys = read(clean, "01_Surveys")
    cmap = {r["legacySurveyId"]: r for r in c_surveys}
    check(set(ids) == set(cmap), "legacySurveyId set differs from final clean workbook")
    DERIVED_COLUMNS = {
        "legacySurveyId", "zipCode", "landArea", "landAreaUnit", "landFrontageFt", "landDepthFt",
        "grossBuildingArea", "numberOfFloors", "advanceAmount", "securityDepositAmount",
        "possessionTerms", "surveyDate", "utilityPhase", "sanctionedLoadKva", "transformerAvailable",
        "transformerCapacityKva", "legalDocumentChecklist", "units",
    }
    UNCHANGED_COLUMNS = [c for c in SURVEY_COLUMNS if c not in DERIVED_COLUMNS]
    diff = Counter(c for r in surveys for c in UNCHANGED_COLUMNS if r[c] != cmap[r["legacySurveyId"]][c])
    check(not diff, f"01_Surveys values differ from final clean unexpectedly: {dict(diff)}")

    # zipCode rules
    zip_stats = Counter()
    for r in surveys:
        lid, z, before = r["legacySurveyId"], r["zipCode"], cmap[r["legacySurveyId"]]["zipCode"]
        if not blank(before):
            check(z == before, f"{lid}: existing zipCode {before!r} was overwritten")
            zip_stats["preserved"] += 1
            continue
        expected_zip = gen.city_zip(r["city"])
        check(z == expected_zip, f"{lid}: zipCode {z!r} != city-level code {expected_zip!r} for {r['city']!r}")
        if not blank(z):
            check(isinstance(z, str) and re.fullmatch(r"\d{5}", z), f"{lid}: zipCode {z!r} is not a 5-digit string")
            zip_stats["inserted"] += 1
        else:
            zip_stats["blank"] += 1

    # possessionTerms / surveyDate must equal the final clean workbook's possession / surveyDate;
    # landArea must equal legacy plotArea whenever plotArea was already populated there
    for r in surveys:
        lid, clean_row = r["legacySurveyId"], cmap[r["legacySurveyId"]]
        check(r["possessionTerms"] == clean_row["possession"], f"{lid}: possessionTerms != legacy possession")
        check(r["surveyDate"] == clean_row["surveyDate"], f"{lid}: surveyDate != legacy surveyDate")
        if not blank(clean_row["plotArea"]):
            check(r["landArea"] == clean_row["plotArea"], f"{lid}: landArea != legacy plotArea")
            check(r["landAreaUnit"] == clean_row["areaUnit"], f"{lid}: landAreaUnit != legacy areaUnit")

    # advanceAmount/securityDepositAmount: only present when the legacy text is a bare number
    commercial_stats = Counter()
    for r in surveys:
        lid, clean_row = r["legacySurveyId"], cmap[r["legacySurveyId"]]
        for src, dst in (("advance", "advanceAmount"), ("security", "securityDepositAmount")):
            raw = clean_row[src]
            val = r[dst]
            if blank(raw):
                check(val is None, f"{lid}: {dst} set without a legacy {src} value")
                continue
            cleaned = re.sub(r"(?i)\b(pkr|rs\.?|rupees)\b", "", str(raw)).replace(",", "").strip()
            if re.fullmatch(r"-?\d+(?:\.\d+)?", cleaned):
                check(val is not None, f"{lid}: {dst} should have parsed from numeric {raw!r}")
                commercial_stats[f"{dst}_mapped"] += 1
            else:
                check(val is None, f"{lid}: {dst} should be null for non-numeric {raw!r}")
                commercial_stats[f"{dst}_unresolved"] += 1

    # ------------------------------------------------ floors
    f_head, floors = data["02_Survey_Floors"]
    survey_ids = set(ids)
    check(all(f["legacySurveyId"] in survey_ids for f in floors), "floor references unknown survey")
    by_survey = defaultdict(list)
    for f in floors:
        by_survey[f["legacySurveyId"]].append(f)
    floor_stats = Counter()
    dup_floor_code_surveys = 0
    for lid, fs in by_survey.items():
        codes = [f["floorCode"] for f in fs if f["floorCode"]]
        if len(codes) != len(set(codes)):
            dup_floor_code_surveys += 1
        for f in fs:
            check(f["floorCode"] is None or FLOOR_CODE_RE.match(f["floorCode"]), f"{lid}: bad floorCode {f['floorCode']!r}")
            for dim in ("floorWidthFt", "floorDepthFt", "floorAreaSqFt"):
                check(f[dim] is None or f[dim] > 0, f"{lid}: {dim} {f[dim]!r} must be > 0 when present")
            if f["floorWidthFt"] is not None and f["floorDepthFt"] is not None:
                calc = f["floorWidthFt"] * f["floorDepthFt"]
                check(f["floorAreaSqFt"] is not None and abs(f["floorAreaSqFt"] - calc) < 1e-6,
                      f"{lid}: floorAreaSqFt {f['floorAreaSqFt']!r} != floorWidthFt*floorDepthFt {calc}")
            if f["floorWidthFt"] is not None and f["floorDepthFt"] is not None and f["floorAreaSqFt"] is not None:
                floor_stats["fully_specified"] += 1
            elif f["floorAreaSqFt"] is not None:
                floor_stats["area_only"] += 1
            else:
                floor_stats["label_only"] += 1
    # gross-area recalculation classification
    gross_class = Counter()
    for lid, fs in by_survey.items():
        physical = [f for f in fs if f["floorCode"] is not None]
        if not physical:
            gross_class["NO_PHYSICAL_FLOORS"] += 1
            continue
        measured = [f for f in physical if f["floorAreaSqFt"] is not None]
        if len(measured) == len(physical):
            gross_class["VERIFIED_OR_RECALCULATED"] += 1
        else:
            gross_class["INSUFFICIENT_FLOOR_DATA"] += 1

    # units[] must contain no leftover floor-shaped objects anywhere in the sheet
    units_floor_leftovers = 0
    for r in surveys:
        if blank(r["units"]):
            continue
        try:
            for item in json.loads(r["units"]):
                if isinstance(item, dict) and "dimensions" in item:
                    units_floor_leftovers += 1
        except (TypeError, ValueError):
            pass
    check(units_floor_leftovers == 0, f"{units_floor_leftovers} units[] entries still look like floor data")

    # ------------------------------------------------ contacts / attachments
    _, contacts = data["03_Survey_Contacts"]
    cids = [c["legacySurveyId"] for c in contacts]
    check(len(contacts) == 7398, f"03_Survey_Contacts rows={len(contacts)}")
    check(len(set(cids)) == len(cids) and set(cids) == survey_ids, "contacts are not one per survey")
    _, c_contacts = read(clean, "02_Survey_Contacts")
    check(contacts == c_contacts, "03_Survey_Contacts differs from final clean")

    _, atts = data["04_Attachments"]
    _, c_atts = read(clean, "03_Attachments")
    check(len(atts) == len(c_atts), f"04_Attachments rows={len(atts)}, final clean has {len(c_atts)}")
    check(atts == c_atts, "04_Attachments differs from final clean")
    check(all(a["legacySurveyId"] in survey_ids for a in atts), "attachment references unknown survey")

    # ------------------------------------------------ seed exceptions
    _, exc = data["05_Seed_Exceptions"]
    check(all(e["entity"] == "Dataset" or e["legacySurveyId"] in survey_ids for e in exc),
          "seed exception references unknown survey")
    ALLOWED_FORBIDDEN_REASONS = {
        "COMMERCIAL_AMOUNT_UNPARSEABLE", "PENDING_BUSINESS_DECISION", "LAND_DIMENSION_ORDER_UNKNOWN",
        "PLOT_DIMENSIONS_UNPARSEABLE", "COVERED_SIZE_MULTI_FLOOR_AMBIGUOUS", "COVERED_SIZE_AREA_MISMATCH",
        "COVERED_SIZE_RECOVERED_FLOOR_DIMENSIONS",
    }
    check(not any(e["field"] in FORBIDDEN and e["reason"] not in ALLOWED_FORBIDDEN_REASONS for e in exc),
          "seed exceptions still reference removed fields outside the audit-trail reasons")

    created_by = [r["createdById"] for r in surveys if not blank(r["createdById"])]
    surveys_units_null = sum(1 for r in surveys if blank(r["units"]))
    exc_by_reason = Counter(e["reason"] for e in exc)
    report = {
        "ok": not errors,
        "errorCount": len(errors),
        "errors": errors[:100],
        "sheets": seed.sheetnames,
        "rows": {n: len(data[n][1]) for n in SHEETS},
        "surveyColumns": len(s_head),
        "removedFieldsAbsent": {f: f not in s_head for f in FORBIDDEN},
        "jsonCellsParsed": dict(json_ok),
        "zipCode": dict(zip_stats),
        "commercialAmounts": dict(commercial_stats),
        "floorStats": dict(floor_stats),
        "floorSurveysWithDuplicateCode": dup_floor_code_surveys,
        "grossAreaClassification": dict(gross_class),
        "unitsNullSurveys": surveys_units_null,
        "unitsGenuineSurveys": len(surveys) - surveys_units_null,
        "coveredSizeRecovered": exc_by_reason["COVERED_SIZE_RECOVERED_FLOOR_DIMENSIONS"],
        "coveredSizeRejectedMultiFloor": exc_by_reason["COVERED_SIZE_MULTI_FLOOR_AMBIGUOUS"],
        "coveredSizeRejectedAreaMismatch": exc_by_reason["COVERED_SIZE_AREA_MISMATCH"],
        "floorNameUnmapped": exc_by_reason["FLOOR_NAME_UNMAPPED"],
        "createdById": {"populated": len(created_by), "blank": len(surveys) - len(created_by),
                        "distinctEmails": len(set(created_by))},
        "seedExceptionsByReason": dict(exc_by_reason.most_common()),
    }
    print(json.dumps(report, indent=2, default=str))
    if errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
