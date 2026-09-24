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
from collections import Counter  # noqa: E402
from datetime import datetime  # noqa: E402
from pathlib import Path  # noqa: E402

from openpyxl import load_workbook  # noqa: E402

HERE = Path(__file__).resolve().parent
SEED = HERE / "cre-backend-seed.xlsx"
CLEAN = HERE / "legacy-survey-final-clean-7398.xlsx"

SHEETS = ["01_Surveys", "02_Survey_Contacts", "03_Attachments", "04_Seed_Exceptions"]
SURVEY_COLUMNS = """legacySurveyId
serialNumber propertyName address locationCoordinates city area zipCode yearBuilt yearRenovated
buildingClass propertyType
ownerName ownershipEntity propertyManager acquisitionDate
plotArea areaUnit plotSize landShape topography zoning
grossBuildingArea rentableArea usableArea coveredSize numberOfFloors ceilingHeight
constructionType exteriorMaterial roofType foundationType parkingSpaces loadingDocks elevatorCount
hvacDetails fireProtection
availabilityStatus propertyStatus availableFor
financialIncome financialExpenses financialMetrics occupancySurvey leaseSurvey utilityDetails units
siteDetails buildingCondition environmentalDetails legalDetails marketDetails attachments
createdById updatedById createdAt updatedAt freshSurveyCount neighboringBusinessIds submittedAt""".split()
CONTACT_COLUMNS = ["legacySurveyId", "name", "role", "phone", "email", "createdAt"]
ATTACHMENT_COLUMNS = ["legacySurveyId", "attachmentType", "originalUrl", "originalFilename", "mimeType",
                      "fileSizeBytes"]
EXCEPTION_COLUMNS = ["entity", "legacySurveyId", "field", "rawValue", "normalizedValue", "reason",
                     "severity", "notes"]
FORBIDDEN = ["askingAmount", "possession", "survey", "advance", "security", "gracePeriod", "increment",
             "agreementPeriod", "surveyDate"]
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
                "buildingCondition", "environmentalDetails", "legalDetails", "marketDetails", "attachments",
                "neighboringBusinessIds"]
EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


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

    data = {n: read(seed, n) for n in SHEETS}
    expected = {"01_Surveys": SURVEY_COLUMNS, "02_Survey_Contacts": CONTACT_COLUMNS,
                "03_Attachments": ATTACHMENT_COLUMNS, "04_Seed_Exceptions": EXCEPTION_COLUMNS}
    for n in SHEETS:
        h = data[n][0]
        check(h == expected[n], f"{n}: columns differ from spec")
        bad = [c for c in h if c in NOT_ALLOWED or (n != "04_Seed_Exceptions" and c in FORBIDDEN)]
        check(not bad, f"{n}: forbidden/helper/system columns {bad}")

    # ------------------------------------------------ surveys
    s_head, surveys = data["01_Surveys"]
    ids = [r["legacySurveyId"] for r in surveys]
    check(len(SURVEY_COLUMNS) == 60 and len(s_head) == 60, f"01_Surveys has {len(s_head)} columns, expected 60")
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

    # every kept value equals the final clean workbook, row by row on legacySurveyId;
    # zipCode is the only column allowed to differ (city-level default added)
    _, c_surveys = read(clean, "01_Surveys")
    cmap = {r["legacySurveyId"]: r for r in c_surveys}
    check(set(ids) == set(cmap), "legacySurveyId set differs from final clean workbook")
    diff = Counter(c for r in surveys for c in SURVEY_COLUMNS
                   if c != "zipCode" and r[c] != cmap[r["legacySurveyId"]][c])
    check(not diff, f"01_Surveys values differ from final clean: {dict(diff)}")

    # zipCode rules
    spec = importlib.util.spec_from_file_location("seedgen", HERE / "generate-cre-backend-seed.py")
    gen = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(gen)
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

    # ------------------------------------------------ contacts / attachments
    survey_ids = set(ids)
    _, contacts = data["02_Survey_Contacts"]
    cids = [c["legacySurveyId"] for c in contacts]
    check(len(contacts) == 7398, f"02_Survey_Contacts rows={len(contacts)}")
    check(len(set(cids)) == len(cids) and set(cids) == survey_ids, "contacts are not one per survey")
    _, c_contacts = read(clean, "02_Survey_Contacts")
    check(contacts == c_contacts, "02_Survey_Contacts differs from final clean")

    _, atts = data["03_Attachments"]
    _, c_atts = read(clean, "03_Attachments")
    check(len(atts) == len(c_atts), f"03_Attachments rows={len(atts)}, final clean has {len(c_atts)}")
    check(atts == c_atts, "03_Attachments differs from final clean")
    check(all(a["legacySurveyId"] in survey_ids for a in atts), "attachment references unknown survey")

    # ------------------------------------------------ seed exceptions
    _, exc = data["04_Seed_Exceptions"]
    check(all(e["entity"] == "Dataset" or e["legacySurveyId"] in survey_ids for e in exc),
          "seed exception references unknown survey")
    check(not any(e["field"] in FORBIDDEN and e["reason"] != "SCHEMA_REMOVE_SURVEY_DATE" for e in exc),
          "seed exceptions still reference removed fields")

    created_by = [r["createdById"] for r in surveys if not blank(r["createdById"])]
    report = {
        "ok": not errors,
        "errorCount": len(errors),
        "errors": errors[:100],
        "sheets": seed.sheetnames,
        "rows": {n: len(data[n][1]) for n in SHEETS},
        "surveyColumns": len(s_head),
        "surveyColumnList": s_head,
        "removedFieldsAbsent": {f: f not in s_head for f in FORBIDDEN},
        "jsonCellsParsed": dict(json_ok),
        "zipCode": dict(zip_stats),
        "createdById": {"populated": len(created_by), "blank": len(surveys) - len(created_by),
                        "distinctEmails": len(set(created_by))},
        "seedExceptionsByReason": dict(Counter(e["reason"] for e in exc).most_common()),
    }
    print(json.dumps(report, indent=2, default=str))
    if errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
