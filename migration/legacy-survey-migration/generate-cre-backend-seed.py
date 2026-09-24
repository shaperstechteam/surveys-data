#!/usr/bin/env python3
"""Generate the CRE backend seed workbook from the final clean workbook.

Usage (from the repo root):

    python migration/legacy-survey-migration/generate-cre-backend-seed.py

Reads (read-only):
    legacy-survey-final-clean-7398.xlsx     the validated historical/reference workbook

Writes:
    cre-backend-seed.xlsx
        01_Surveys          7,398 rows, aligned with the current Survey model
        02_Survey_Floors    one row per legacy units[] floor entry (SurveyFloor)
        03_Survey_Contacts  copied unchanged
        04_Attachments      copied unchanged
        05_Seed_Exceptions  only the blockers/warnings that matter when seeding CRE

askingAmount, possession, survey, advance, security, gracePeriod, increment and
agreementPeriod stay on legacy-survey-final-clean-7398.xlsx as the historical record.
Of those: possession becomes possessionTerms (direct copy); advance/security become
advanceAmount/securityDepositAmount only where the raw text is an actual numeric amount
(most are "N Months" of rent, not a currency figure, and are left null + flagged rather
than guessed); askingAmount/gracePeriod/increment/agreementPeriod have no approved Survey
destination yet and are reported as dataset-level PENDING_BUSINESS_DECISION exceptions.
surveyDate is restored (final clean already nulls unparseable legacy dates).

The legacy units[] column actually holds floor-level data ({floor, squareFeet?,
dimensions?} — built that way by generate-final-clean-workbook.py), not tenant/unit
records. Every units[] entry is reclassified here: floor-shaped entries become
02_Survey_Floors rows and are dropped from units; only entries carrying genuine
tenant/unit signals (tenant, number, leaseStart/End, monthlyRent, annualRent,
occupancyStatus) stay in units. Since the legacy dataset never carries those signals,
01_Surveys.units is null for every row in this seed - nothing is fabricated to fill it.

plotSize, coveredSize and ceilingHeight are dropped from the seed (still on final
clean for audit). plotSize's native area value becomes landArea/landAreaUnit; its
raw dimension pair, when present, is preserved via a LAND_DIMENSION_ORDER_UNKNOWN
exception rather than being assigned to landFrontageFt/landDepthFt, since the legacy
column never establishes which number is frontage and which is depth. coveredSize is
used only as a last-resort SurveyFloor dimension recovery for single-floor surveys
whose stated area already matches the covered-size calculation; ambiguous
(multi-floor) or mismatched cases are left null and flagged instead. Nothing is
imported or uploaded.
"""
from __future__ import annotations

import json
import re
import sys

sys.dont_write_bytecode = True

from collections import Counter  # noqa: E402
from datetime import datetime  # noqa: E402
from decimal import Decimal, InvalidOperation  # noqa: E402
from pathlib import Path  # noqa: E402

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

HERE = Path(__file__).resolve().parent
SOURCE = HERE / "legacy-survey-final-clean-7398.xlsx"
OUTPUT = HERE / "cre-backend-seed.xlsx"

# Kept in legacy-survey-final-clean-7398.xlsx only; no longer copied verbatim into the seed.
REMOVED_FIELDS = ["askingAmount", "possession", "survey", "advance", "security", "gracePeriod",
                  "increment", "agreementPeriod"]
# Have no approved Survey destination this round; reported as dataset-level exceptions instead.
NO_DESTINATION_FIELDS = ["askingAmount", "gracePeriod", "increment", "agreementPeriod"]

SEED_SURVEY_FIELDS = [
    "serialNumber", "propertyName", "address", "locationCoordinates", "city", "area", "zipCode",
    "yearBuilt", "yearRenovated", "buildingClass", "propertyType",
    "ownerName", "ownershipEntity", "propertyManager", "acquisitionDate",
    "landArea", "landAreaUnit", "landFrontageFt", "landDepthFt", "landShape", "topography", "zoning",
    "grossBuildingArea", "rentableArea", "usableArea", "numberOfFloors",
    "constructionType", "exteriorMaterial", "roofType", "foundationType", "parkingSpaces",
    "loadingDocks", "elevatorCount", "hvacDetails", "fireProtection",
    "utilityPhase", "sanctionedLoadKva", "transformerAvailable", "transformerCapacityKva",
    "availabilityStatus", "propertyStatus", "availableFor",
    "advanceAmount", "securityDepositAmount", "possessionTerms",
    "surveyDate",
    "financialIncome", "financialExpenses", "financialMetrics", "occupancySurvey", "leaseSurvey",
    "utilityDetails", "units", "siteDetails", "buildingCondition", "environmentalDetails",
    "legalDetails", "legalDocumentChecklist", "marketDetails", "attachments",
    "createdById", "updatedById", "createdAt", "updatedAt", "freshSurveyCount",
    "neighboringBusinessIds", "submittedAt",
]
SURVEY_HEADERS = ["legacySurveyId"] + SEED_SURVEY_FIELDS

# NOTE: the actual Prisma SurveyFloor model may still name these lengthFt/widthFt/areaSqFt.
# This workbook uses the CRE-oriented names requested (floorWidthFt/floorDepthFt/floorAreaSqFt);
# reconcile with the backend schema before import if it hasn't been renamed there too.
FLOOR_HEADERS = ["legacySurveyId", "floorCode", "floorName", "sortOrder", "floorWidthFt", "floorDepthFt",
                 "floorAreaSqFt", "rentableAreaSqFt", "usableAreaSqFt", "ceilingHeightFt", "floorUse",
                 "notes", "attachments"]

# keys that indicate a units[] entry is a genuine tenant/unit record, not floor-level data
UNIT_SIGNAL_KEYS = {"tenant", "number", "leaseStart", "leaseEnd", "monthlyRent", "annualRent",
                    "occupancyStatus"}
UNIT_OUTPUT_KEYS = ["number", "tenant", "floor", "squareFeet", "leaseStart", "leaseEnd",
                    "monthlyRent", "annualRent", "occupancyStatus"]

EXCEPTION_HEADERS = ["entity", "legacySurveyId", "field", "rawValue", "normalizedValue",
                     "reason", "severity", "notes"]

# City-level DEFAULT postal code (main GPO / main delivery office), not the exact
# neighbourhood code. Source: Pakistan Post "Post Code Directory of Delivery Post
# Offices" (pakpost.gov.pk/images/national post code directory.pdf); the office
# named here is the directory entry the code was taken from. Keys are matched after
# trimming and case-folding the city; the city text itself is never changed.
# Deliberately unmapped: Swat (a district, not one city), rawat and Mumtaz City
# (no matching directory entry).
CITY_ZIP = {
    "lahore": ("54000", "LAHORE GPO"),
    "islamabad": ("44000", "ISLAMABAD GPO"),
    "rawalpindi": ("46000", "RAWALPINDI GPO"),
    "karachi": ("74200", "KARACHI GPO"),
    "faisalabad": ("38000", "FAISALABAD GPO"),
    "peshawar": ("25000", "PESHAWAR GPO"),
    "jhelum": ("49600", "JHELUM GPO"),
    "gujranwala": ("52250", "GUJRANWALA GPO"),
    "gujrat": ("50700", "GUJRAT GPO"),
    "sialkot": ("51310", "SIALKOT GPO"),
    "multan": ("60000", "MULTAN GPO"),
    "mirpur ajk": ("10250", "MIRPUR GPO (Azad Kashmir)"),
    "abbottabad": ("22010", "ABBOTTABAD GPO"),
    "sargodha": ("40100", "SARGODHA GPO"),
    "wah cantt.": ("47040", "WAH CANTT. GPO"),
    "mardan": ("23200", "MARDAN GPO"),
    "hyderabad": ("71000", "HYDERABAD GPO"),
    "mandi bahauddin": ("50400", "MANDI BAHAUDDIN GPO"),
    "kasur": ("55050", "KASUR GPO"),
    "rahim yar khan": ("64200", "RAHIMYAR KHAN GPO"),
    "sahiwal": ("57000", "SAHIWAL GPO"),
    "sukkur": ("65200", "SUKKUR GPO"),
    "murree": ("47150", "MURREE GPO"),
    "attock": ("43600", "ATTOCK GPO"),
    "sheikhupura": ("39350", "QILA SHEIKHUPURA GPO"),
    "okara": ("56300", "OKARA GPO"),
    "haripur": ("22620", "HARIPUR GPO"),
    "bahawalpur": ("63100", "BAHAWALPUR GPO"),
    "quetta": ("87300", "QUETTA GPO"),
    "muzaffarabad": ("13100", "MUZAFFARABAD GPO"),
    "dera ghazi khan": ("32200", "DERA GHAZI KHAN GPO"),
    "daska": ("51010", "DASKA"),
    "narowal": ("51600", "NAROWAL GPO"),
    "vehari": ("61100", "VEHARI GPO"),
    "burewala": ("61010", "BUREWALA"),
    "chakwal": ("48800", "CHAKWAL GPO"),
    "hafizabad": ("52110", "HAFIZ ABAD"),
    "wazirabad": ("52000", "WAZIRABAD"),
    "jhang": ("35200", "JHANG GPO"),
    "taxila": ("47080", "TAXILA"),
    "wah": ("47000", "WAH"),
    "mianwali": ("42200", "MIANWALI GPO"),
    "fateh jang": ("43350", "FATEH JANG"),
    "muridke": ("39000", "MURIDKE"),
    "kharian": ("50090", "KHARIAN CITY"),
    "lalamusa": ("50200", "LALA MUSA"),
    "bahawalnagar": ("62300", "BAHAWAL NAGAR GPO"),
    "lodhran": ("59320", "LODHRAN"),
    "toba tek singh": ("36050", "TOBA TAKE SINGH GPO"),
    "khanewal": ("58150", "KHANEWAL GPO"),
    "gujar khan": ("47850", "GUJAR KHAN GPO"),
    "mirpur khas": ("69000", "MIRPUR KHAS GPO"),
    "dina": ("49400", "DINA"),
    "mansehra": ("21300", "MANSEHRA GPO"),
    "layyah": ("31200", "LAYYAH GPO"),
    "sarai alamgir": ("50000", "SARAI ALAMGIR"),
    "chiniot": ("35400", "CHINIOT"),
    "pir mahal": ("36300", "PIR MAHAL"),
    "dara adam khel": ("26100", "DARA ADAM KHEL"),
    "hazro": ("43440", "HAZRO"),
    "arifwala": ("57450", "ARIF WALA"),
    "dinga": ("50280", "DINGA"),
    "bannu": ("28100", "BANNU GPO"),
    "phalia": ("50430", "PHALIA"),
    "joharabad": ("41200", "JAUHAR ABAD"),
    "hattar": ("43394", "HATTAR"),
    "nowshera": ("24100", "NOWSHERA GPO"),
    "ali pur chatha": ("52080", "ALI PUR CHATHA"),
}


def city_zip(city):
    if city is None or not str(city).strip():
        return None
    hit = CITY_ZIP.get(str(city).strip().casefold())
    return hit[0] if hit else None


# ---------------------------------------------------------------- floor names

ORDINALS = ["First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth",
            "Tenth", "Eleventh", "Twelfth", "Thirteenth", "Fourteenth", "Fifteenth", "Sixteenth",
            "Seventeenth", "Eighteenth", "Nineteenth", "Twentieth"]

FLOOR_CODES = {
    "basement": "B1",
    "basement 1": "B1",
    "basement 2": "B2",
    "lower ground floor": "LG",
    "lower ground": "LG",
    "upper ground floor": "UG",
    "upper ground": "UG",
    "ground floor": "G",
    "ground": "G",
    "mezzanine floor": "M",
    "mezzanine": "M",
    "rooftop": "RF",
    "roof top": "RF",
}
for _i, _name in enumerate(ORDINALS, start=1):
    FLOOR_CODES[f"{_name.lower()} floor"] = f"{_i}F"

# known legacy typos/spacing quirks that would otherwise miss FLOOR_CODES
FLOOR_TYPO_FIXES = {
    "forth floor": "fourth floor",
    "second floord": "second floor",
    "twelveth floor": "twelfth floor",
}

FLOOR_SHORT_ORDINAL_RE = re.compile(r"^(\d+)(?:st|nd|rd|th)\s+floor$")


def normalize_floor_name(raw):
    v = re.sub(r"\s+", " ", str(raw).strip()).lower()
    if not v:
        return None
    v = FLOOR_TYPO_FIXES.get(v, v)
    code = FLOOR_CODES.get(v)
    if code:
        return code
    m = FLOOR_SHORT_ORDINAL_RE.match(v)
    if m:
        return f"{m.group(1)}F"
    return None


# ---------------------------------------------------------------- dimensions

FLOOR_DIM_RE = re.compile(r"^(\d+(?:\.\d+)?)\s*ft\s*x\s*(\d+(?:\.\d+)?)\s*ft\.?$", re.I)
PLOT_DIM_RE = re.compile(r"^(\d+(?:\.\d+)?)\s*(?:ft)?\s*['′]?\s*[xX*×]\s*"
                         r"(\d+(?:\.\d+)?)\s*(?:ft)?\s*['′]?\.?$")
AREA_UNIT_RE = re.compile(
    r"^\d+(?:\.\d+)?\s*(kanal|kanals|marla|marlas|marls|acre|acres|sq\.?\s*y(?:ar)?ds?|gazz?|"
    r"sq\.?\s*f(?:ee)?t|sft|sqft)\s*\.?$", re.I)


def parse_dimensions(raw, pattern):
    if not raw:
        return None, None
    m = pattern.match(str(raw).strip())
    if not m:
        return None, None
    return float(m.group(1)), float(m.group(2))


# ---------------------------------------------------------------- commercial amounts

def parse_commercial_amount(raw):
    """Return (Decimal amount or None, unparsed reason or None)."""
    v = str(raw).strip()
    if not v:
        return None, None
    cleaned = re.sub(r"(?i)\b(pkr|rs\.?|rupees)\b", "", v).replace(",", "").strip()
    if re.fullmatch(r"-?\d+(?:\.\d+)?", cleaned):
        try:
            return Decimal(cleaned), None
        except InvalidOperation:
            pass
    m = re.match(r"^(\d+(?:\.\d+)?)\s*months?$", v, re.I)
    if m:
        return None, f"{m.group(1)} months of rent, not a currency amount"
    return None, "unparseable as a currency amount"


EXCEPTION_HEADERS_SET = set(EXCEPTION_HEADERS)

# per-row exceptions from the final clean workbook that still matter for the seed
KEEP_ROW_REASONS = {
    "INVALID_LEGACY_DATE", "BUSINESS_DEVELOPER_MISSING", "MISSING_COORDINATES", "INVALID_COORDINATES",
    "COORDINATES_OUTSIDE_PAKISTAN", "CONFLICT_WITH_POSSESSION", "DUPLICATE_CANDIDATE",
    "UNIT_FLOOR_LABEL_MISSING", "UNPARSEABLE_FLOOR_AREA",
    "REVIEW_CP_SEMANTICS", "INVALID_NAME", "REVIEW_MULTI_PHONE",          # former HOLD contacts
}
KEEP_DATASET_REASONS = {
    "CREATED_BY_ID_REQUIRES_USER_LOOKUP", "TIMESTAMP_TIMEZONE_UNCONFIRMED", "LEGACY_STATUS_CONSTANT_ACTIVE",
}


def read_sheet(wb, name):
    it = wb[name].iter_rows(values_only=True)
    headers = list(next(it))
    return headers, [dict(zip(headers, r)) for r in it if any(v is not None for v in r)]


def build():
    src = load_workbook(SOURCE, read_only=True)
    s_head, surveys = read_sheet(src, "01_Surveys")
    c_head, contacts = read_sheet(src, "02_Survey_Contacts")
    a_head, attachments = read_sheet(src, "03_Attachments")
    _, exceptions = read_sheet(src, "04_Exceptions")

    NEW_FIELDS = ("landArea", "landAreaUnit", "landFrontageFt", "landDepthFt", "utilityPhase",
                 "sanctionedLoadKva", "transformerAvailable", "transformerCapacityKva", "advanceAmount",
                 "securityDepositAmount", "possessionTerms", "legalDocumentChecklist")
    required = ["legacySurveyId"] + [f for f in SEED_SURVEY_FIELDS if f not in NEW_FIELDS]
    missing = [f for f in required if f not in s_head]
    if missing:
        raise SystemExit(f"final clean workbook lacks {missing}")
    if len(surveys) != 7398:
        raise SystemExit(f"expected 7,398 surveys, found {len(surveys)}")

    floor_rows = []
    commercial_stats = Counter()
    floor_stats = Counter()
    unit_stats = Counter()
    covered_size_stats = Counter()
    seed_exc = []

    def exc(entity, lid, field, raw, normalized, reason, severity, notes):
        seed_exc.append({"entity": entity, "legacySurveyId": lid, "field": field, "rawValue": raw,
                         "normalizedValue": normalized, "reason": reason, "severity": severity,
                         "notes": notes})

    seed_surveys = []
    for r in surveys:
        lid = r["legacySurveyId"]
        row = {h: r[h] for h in SEED_SURVEY_FIELDS if h in r}
        row["legacySurveyId"] = lid

        # zipCode: keep any existing value; otherwise the city-level default (as text)
        if row["zipCode"] in (None, ""):
            row["zipCode"] = city_zip(row["city"])

        # -------------------------------------------------- land dimensions (was plotSize/plotArea/areaUnit)
        row["landArea"] = r.get("plotArea")
        row["landAreaUnit"] = r.get("areaUnit")
        row["landFrontageFt"] = None
        row["landDepthFt"] = None
        raw_plot_size = r.get("plotSize")
        length, depth = parse_dimensions(raw_plot_size, PLOT_DIM_RE)
        if length is not None:
            computed_area = length * depth
            if row["landArea"] in (None, ""):
                row["landArea"] = computed_area
                if row["landAreaUnit"] in (None, ""):
                    row["landAreaUnit"] = "Sq Ft"
            exc("Survey", lid, "plotSize", raw_plot_size, computed_area, "LAND_DIMENSION_ORDER_UNKNOWN",
                "LOW", "Legacy plotSize does not establish which number is frontage vs. depth; area is "
                "safe to compute (commutative) but landFrontageFt/landDepthFt are left null")
        elif raw_plot_size and not AREA_UNIT_RE.match(str(raw_plot_size).strip()):
            exc("Survey", lid, "plotSize", raw_plot_size, None, "PLOT_DIMENSIONS_UNPARSEABLE", "LOW",
                "Neither a recognizable dimension pair nor a known area-unit expression")

        # -------------------------------------------------- commercial terms
        possession_raw = r.get("possession")
        row["possessionTerms"] = possession_raw

        for src_field, dst_field, stat_key in (("advance", "advanceAmount", "advance"),
                                                ("security", "securityDepositAmount", "security")):
            raw_val = r.get(src_field)
            if raw_val in (None, ""):
                row[dst_field] = None
                continue
            amount, reason = parse_commercial_amount(raw_val)
            row[dst_field] = float(amount) if amount is not None else None
            if amount is not None:
                commercial_stats[f"{stat_key}_mapped"] += 1
            else:
                commercial_stats[f"{stat_key}_unresolved"] += 1
                exc("Survey", lid, src_field, raw_val, None, "COMMERCIAL_AMOUNT_UNPARSEABLE", "LOW", reason)

        # -------------------------------------------------- surveyDate (restored)
        row["surveyDate"] = r.get("surveyDate")
        if row["surveyDate"] is not None:
            commercial_stats["surveyDate_recovered"] += 1

        # -------------------------------------------------- utilities / legal checklist (no legacy source)
        row["utilityPhase"] = None
        row["sanctionedLoadKva"] = None
        row["transformerAvailable"] = None
        row["transformerCapacityKva"] = None
        row["legalDocumentChecklist"] = None

        # -------------------------------------------------- units[] reclassification: floor-data vs. genuine units
        units_raw = r.get("units")
        floor_items, genuine_units = [], []
        if units_raw:
            try:
                items = json.loads(units_raw)
            except (TypeError, ValueError):
                items = []
            for item in items:
                if isinstance(item, dict) and UNIT_SIGNAL_KEYS & item.keys():
                    genuine_units.append({k: item.get(k) for k in UNIT_OUTPUT_KEYS if k in item})
                else:
                    floor_items.append(item)

        if genuine_units:
            row["units"] = json.dumps(genuine_units)
            unit_stats["surveys_with_genuine_units"] += 1
        else:
            row["units"] = None
            unit_stats["surveys_units_became_null"] += 1
        unit_stats["floor_objects_removed_from_units"] += len(floor_items)

        # -------------------------------------------------- floors (from the reclassified floor-only entries)
        survey_floors = []
        for sort_order, item in enumerate(floor_items):
            floor_name = item.get("floor")
            dims = item.get("dimensions")
            stated_sqft = item.get("squareFeet")
            width_ft, depth_ft = parse_dimensions(dims, FLOOR_DIM_RE)
            calc_area = width_ft * depth_ft if width_ft is not None else None
            area_sqft = None
            notes = None
            if calc_area is not None:
                area_sqft = calc_area
                if isinstance(stated_sqft, (int, float)) and abs(calc_area - stated_sqft) > 1:
                    notes = (f"legacy stated squareFeet={stated_sqft}; using calculated "
                             f"floorWidthFt*floorDepthFt={calc_area:g}")
                    exc("SurveyFloor", lid, "floorAreaSqFt", stated_sqft, calc_area,
                        "FLOOR_AREA_MISMATCH", "MEDIUM", notes)
                    floor_stats["mismatch"] += 1
            elif isinstance(stated_sqft, (int, float)):
                area_sqft = float(stated_sqft)

            floor_code = normalize_floor_name(floor_name) if floor_name else None
            if floor_name and floor_code is None:
                exc("SurveyFloor", lid, "floorName", floor_name, None, "FLOOR_NAME_UNMAPPED", "LOW",
                    "No confident floor-code mapping; kept as floorName only")

            if width_ft is not None:
                floor_stats["fully_specified"] += 1
            elif area_sqft is not None:
                floor_stats["area_only"] += 1
            else:
                floor_stats["label_only"] += 1

            survey_floors.append({
                "legacySurveyId": lid, "floorCode": floor_code, "floorName": floor_name,
                "sortOrder": sort_order, "floorWidthFt": width_ft, "floorDepthFt": depth_ft,
                "floorAreaSqFt": area_sqft, "rentableAreaSqFt": None, "usableAreaSqFt": None,
                "ceilingHeightFt": None, "floorUse": None, "notes": notes, "attachments": None,
            })

        # -------------------------------------------------- coveredSize: single-floor dimension recovery only
        raw_covered_size = r.get("coveredSize")
        if raw_covered_size:
            cs_width, cs_depth = parse_dimensions(raw_covered_size, FLOOR_DIM_RE)
            physical_now = [f for f in survey_floors if f["floorCode"] is not None]
            if cs_width is None:
                covered_size_stats["unparseable"] += 1
            elif len(physical_now) != 1:
                covered_size_stats["rejected_multi_floor_ambiguous"] += 1
                exc("SurveyFloor", lid, "coveredSize", raw_covered_size, None,
                    "COVERED_SIZE_MULTI_FLOOR_AMBIGUOUS", "LOW",
                    f"{len(physical_now)} physical floor(s); relationship between coveredSize and a "
                    "specific floor is not deterministic")
            else:
                only = physical_now[0]
                cs_area = cs_width * cs_depth
                if only["floorAreaSqFt"] is not None and abs(cs_area - only["floorAreaSqFt"]) <= 1:
                    only["floorWidthFt"], only["floorDepthFt"] = cs_width, cs_depth
                    covered_size_stats["recovered"] += 1
                    exc("SurveyFloor", lid, "coveredSize", raw_covered_size, cs_area,
                        "COVERED_SIZE_RECOVERED_FLOOR_DIMENSIONS", "LOW",
                        f"Single physical floor; coveredSize area matches floorAreaSqFt={only['floorAreaSqFt']:g}")
                else:
                    covered_size_stats["rejected_area_mismatch"] += 1
                    exc("SurveyFloor", lid, "coveredSize", raw_covered_size, cs_area,
                        "COVERED_SIZE_AREA_MISMATCH", "LOW",
                        f"coveredSize area {cs_area:g} does not match floorAreaSqFt="
                        f"{only['floorAreaSqFt']!r}; dimensions left null")

        floor_rows.extend(survey_floors)

        # -------------------------------------------------- building totals
        physical = [f for f in survey_floors if f["floorCode"] is not None]
        if physical:
            measured = [f for f in physical if f["floorAreaSqFt"] is not None]
            if len(measured) == len(physical):
                new_gross = sum(f["floorAreaSqFt"] for f in measured)
                new_count = len(physical)
                legacy_gross = row.get("grossBuildingArea")
                if legacy_gross not in (None, "") and abs(float(legacy_gross) - new_gross) > 1:
                    exc("Survey", lid, "grossBuildingArea", legacy_gross, new_gross,
                        "GROSS_AREA_MISMATCH", "MEDIUM",
                        f"Recalculated from {new_count} SurveyFloor rows; legacy value replaced")
                row["grossBuildingArea"] = new_gross
                row["numberOfFloors"] = new_count
            else:
                partial_sum = sum(f["floorAreaSqFt"] for f in measured)
                exc("Survey", lid, "grossBuildingArea", row.get("grossBuildingArea"), None,
                    "FLOOR_TOTAL_INCOMPLETE", "LOW",
                    f"{len(measured)}/{len(physical)} floors have a measured area (partial sum "
                    f"{partial_sum:g}); legacy grossBuildingArea/numberOfFloors kept as-is")

        seed_surveys.append(row)

    # ------------------------------------------------ carried-over seed exceptions
    for e in exceptions:
        if e["entity"] == "Dataset" and e["reason"] in KEEP_DATASET_REASONS:
            seed_exc.append(dict(e))
        elif e["entity"] in ("Survey", "SurveyContact") and e["reason"] in KEEP_ROW_REASONS:
            row = dict(e)
            if row["field"] == "surveyDate/createdAt":
                row["field"] = "createdAt"
                row["notes"] = "No valid date_created; createdAt is blank in the seed"
            seed_exc.append(row)

    reasons = Counter(e["reason"] for e in exceptions)
    blank_type = sum(1 for r in seed_surveys if r["propertyType"] in (None, ""))
    att_excluded = sum(1 for e in exceptions if e["entity"] == "Attachment")

    def dataset(field, raw, reason, severity, notes):
        seed_exc.append({"entity": "Dataset", "legacySurveyId": None, "field": field, "rawValue": raw,
                         "normalizedValue": None, "reason": reason, "severity": severity, "notes": notes})

    for field in NO_DESTINATION_FIELDS:
        count = sum(1 for r in surveys if r.get(field) not in (None, ""))
        if count:
            dataset(field, str(count), "PENDING_BUSINESS_DECISION", "MEDIUM",
                    f"{count:,} surveys have a legacy {field} value; no approved Survey destination "
                    "this round. Preserved in legacy-survey-final-clean-7398.xlsx")

    dataset("area, availableFor, landArea, landAreaUnit, landFrontageFt, landDepthFt", None,
            "SCHEMA_ADD_SURVEY_FIELDS", "HIGH", "These seed columns must exist on the CRE Survey model "
            "before import")
    dataset("locationCoordinates", None, "SCHEMA_LOCATION_COORDINATES_JSON", "HIGH",
            'Values are JSON {"latitude":…,"longitude":…}; CRE column must accept Json, not Decimal(9,6)')
    dataset("units", str(unit_stats["surveys_with_genuine_units"]), "SCHEMA_UNITS_RECLASSIFIED", "LOW",
            f"Legacy units[] held floor-level data, not tenant/unit records; "
            f"{unit_stats['floor_objects_removed_from_units']:,} floor objects moved to SurveyFloor and "
            f"removed from units. Only {unit_stats['surveys_with_genuine_units']:,} surveys have genuine "
            "unit-shaped data (none currently, given the legacy source); the rest are null")
    dataset("createdAt", str(reasons["INVALID_LEGACY_DATE"]), "CREATED_AT_MISSING", "HIGH",
            f"{reasons['INVALID_LEGACY_DATE']:,} surveys have no valid legacy timestamp (rows below); "
            "decide: import with createdAt = import time, or hold them")
    dataset("propertyType", str(blank_type), "PROPERTY_TYPE_UNRESOLVED", "MEDIUM",
            f"{blank_type:,} surveys have blank propertyType (legacy 'building' or blank has no safe enum)")
    dataset("phone", str(reasons["CONTACT_PHONE_INCOMPLETE_OR_INVALID"]), "CONTACT_PHONE_RAW", "MEDIUM",
            f"Contact phones are raw legacy text; {reasons['CONTACT_PHONE_INCOMPLETE_OR_INVALID']:,} are "
            f"incomplete/invalid and {reasons['CONTACT_MULTIPLE_PHONES']:,} hold several numbers. "
            "Importer must normalise or accept as-is")
    unmapped = Counter(str(r["city"]).strip() for r in seed_surveys
                       if r["zipCode"] is None and r["city"] not in (None, ""))
    no_city = sum(1 for r in seed_surveys if r["city"] in (None, ""))
    dataset("zipCode", None, "ZIPCODE_CITY_LEVEL_DEFAULT", "LOW",
            "zipCode is the city's main GPO/delivery-office code from the Pakistan Post directory, not the "
            f"exact neighbourhood code. Left blank: {no_city} surveys with no city; unmapped cities "
            + ", ".join(f"{c} ({n})" for c, n in unmapped.most_common()))
    dataset("attachments", str(att_excluded), "ATTACHMENTS_NOT_IN_MANIFEST", "MEDIUM",
            f"{att_excluded:,} legacy file references are not in 03_Attachments (placeholder, unverified, "
            "unsupported or unavailable). Manifest files still need upload to S3 before AttachmentRecords exist")
    dataset("utilityPhase, sanctionedLoadKva, transformerAvailable, transformerCapacityKva", None,
            "SCHEMA_NO_LEGACY_UTILITY_DATA", "LOW",
            "No legacy source found anywhere in the final clean workbook; left null for every row")
    dataset("legalDocumentChecklist", None, "SCHEMA_NO_LEGACY_CHECKLIST_DATA", "LOW",
            "No structured legacy checklist exists; left null (not NOT_COLLECTED) for every row")

    order = {"Dataset": 0, "Survey": 1, "SurveyFloor": 2, "SurveyContact": 3}
    seed_exc.sort(key=lambda e: (order[e["entity"]], e["legacySurveyId"] or 0, e["field"] or "", e["reason"]))

    write(seed_surveys, floor_rows, c_head, contacts, a_head, attachments, seed_exc)
    print(f"Wrote {OUTPUT.name}: surveys={len(seed_surveys)} x {len(SURVEY_HEADERS)} cols, "
          f"floors={len(floor_rows)}, contacts={len(contacts)}, attachments={len(attachments)}, "
          f"seedExceptions={len(seed_exc)}")
    print("floor stats:", dict(floor_stats))
    print("unit reclassification stats:", dict(unit_stats))
    print("coveredSize recovery stats:", dict(covered_size_stats))
    print("commercial stats:", dict(commercial_stats))
    print("seed exceptions by reason:", dict(Counter(e["reason"] for e in seed_exc).most_common()))


# ---------------------------------------------------------------- write

HEADER_FONT = Font(bold=True)
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
TOP = Alignment(vertical="top")
WRAP = Alignment(vertical="top", wrap_text=True)


def write_sheet(wb, title, headers, records, widths=None, wrap=()):
    ws = wb.create_sheet(title)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    wrap = set(wrap)
    for r, rec in enumerate(records, 2):
        for c, h in enumerate(headers, 1):
            val = rec.get(h)
            if val == "":
                val = None
            cell = ws.cell(r, c)
            cell.value = val
            if isinstance(val, str):
                cell.data_type = "s"
            elif isinstance(val, datetime):
                cell.number_format = "yyyy-mm-dd hh:mm:ss"
            cell.alignment = WRAP if h in wrap else TOP
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(records) + 1}"
    for c, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(c)].width = (widths or {}).get(h, max(12, min(len(h) + 4, 24)))
    ws.sheet_view.zoomScale = 90


def write(surveys, floors, c_head, contacts, a_head, attachments, seed_exc):
    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "01_Surveys", SURVEY_HEADERS, surveys,
                widths={"address": 32, "locationCoordinates": 34, "area": 28, "ownerName": 26,
                        "units": 48, "marketDetails": 48, "createdById": 30, "createdAt": 20,
                        "possessionTerms": 40},
                wrap={"address", "locationCoordinates", "units", "marketDetails", "possessionTerms"})
    write_sheet(wb, "02_Survey_Floors", FLOOR_HEADERS, floors,
                widths={"floorName": 22, "notes": 60})
    write_sheet(wb, "03_Survey_Contacts", c_head, contacts,
                widths={"name": 28, "role": 24, "phone": 28, "createdAt": 20})
    write_sheet(wb, "04_Attachments", a_head, attachments,
                widths={"originalUrl": 70, "originalFilename": 36, "mimeType": 16})
    write_sheet(wb, "05_Seed_Exceptions", EXCEPTION_HEADERS, seed_exc,
                widths={"entity": 14, "field": 26, "rawValue": 50, "normalizedValue": 30,
                        "reason": 34, "severity": 10, "notes": 60},
                wrap={"rawValue", "normalizedValue", "notes"})
    wb.save(OUTPUT)


if __name__ == "__main__":
    build()
