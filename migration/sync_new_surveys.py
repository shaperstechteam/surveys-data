#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sync_new_surveys.py  --  READ-ONLY HostGator (Gravity Forms form 2) -> historical-surveys workbook

What it does
  1. Reads the baseline workbook (never modified) and collects the legacySurveyIds it already has.
  2. Reads Gravity Forms entries for form 2 from HostGator (MySQL SELECT only) or from CSV exports.
  3. Pre-flight report: how many entries are NEW (id not in the baseline), their ID range, and which of
     them cannot be mapped reliably.  --dry-run stops here.
  4. Applies the same transformation rules the existing workbook was built with (reverse-engineered and
     checked against the 7,398 existing rows; see --revalidate) and writes a NEW workbook:
        baseline rows (untouched)  +  new rows appended, same 5 sheets, same columns, same formatting.
  5. Writes a short text/JSON summary: new surveys, exceptions, media needing attention.

What it never does
  * No INSERT/UPDATE/DELETE/DDL on any database (the MySQL session is opened read-only and only SELECTs run).
  * Does not touch the CRE database, does not create surveyors/users, does not download media.
  * Does not modify or overwrite the baseline workbook (it refuses to write to the same path).
  * Credentials are read from environment variables / a hidden prompt; they are never printed or written.
  * Media: only URLs are recorded. fileSizeBytes is left empty, unless you pass --check-urls, which sends
    HTTP HEAD requests (headers only, no file body) to read Content-Length / status.

Usage examples
  # A) CSV exports from phpMyAdmin (queries are in README):  no database credentials needed at all
  python3 sync_new_surveys.py --baseline historical-surveys.xlsx --entry-csv new_entries.csv \
          --meta-csv new_entry_meta.csv --out historical-surveys_updated.xlsx

  # B) Direct MySQL (run on the HostGator account, or remotely if Remote MySQL is already allowed)
  export GF_DB_HOST=localhost GF_DB_NAME=siugplte_websportal GF_DB_USER=...   # password is prompted
  python3 sync_new_surveys.py --baseline historical-surveys.xlsx --mysql --out historical-surveys_updated.xlsx

  # Pre-flight only:
  python3 sync_new_surveys.py --baseline historical-surveys.xlsx --mysql --dry-run

  # C) ALL media of ALL surveys (new surveys are added; every survey's 04_Attachments rows are rebuilt from the dump:
  #    all images + videos of fields 83/103/129, replaced/deleted files updated, unchanged rows kept as they are)
  python3 sync_new_surveys.py --baseline oct-surveys-dump/historical-surveys.xlsx --dump-dir oct-surveys-dump \
          --all-media --check-urls --dry-run
  python3 sync_new_surveys.py --baseline oct-surveys-dump/historical-surveys.xlsx --dump-dir oct-surveys-dump \
          --all-media --check-urls --out historical-surveys_all-media.xlsx
  --check-urls HEADs only the NEW files (no cookies, 8 at a time) to fill their size and leave out dead links.
  Also written next to the output: <out>_summary.json and <out>_attachment_changes.tsv (every added / removed /
  not-added / over-the-app-size-limit file).
"""
import argparse, collections, copy, csv, datetime, getpass, json, os, re, sys, urllib.parse

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required:  pip install openpyxl")

csv.field_size_limit(2 ** 31 - 1)
FORM_ID = 2

# ----------------------------------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------------------------------
JUNK = {'.', 'nil', '-', 'NULL', 'null', 'none', 'None', 'N/A', 'n/a', 'na', 'NA', 'Select', 'Area Not Listed'}
CITYFIX = {'huripur': 'Haripur', 'Muree': 'Murree', 'jhalum': 'Jhelum', 'Burawala': 'Burewala', 'Jehlum': 'Jhelum'}

# city -> zip : exactly the pairs present in the baseline workbook (city-level default zip)
ZIP = {'Jhelum': '49600', 'Dera Ghazi Khan': '32200', 'Lahore': '54000', 'Islamabad': '44000', 'Hafizabad': '52110',
       'Peshawar': '25000', 'Rawalpindi': '46000', 'Abbottabad': '22010', 'Gujrat': '50700', 'Sargodha': '40100',
       'Gujranwala': '52250', 'Multan': '60000', 'Mandi Bahauddin': '50400', 'Okara': '56300', 'Faisalabad': '38000',
       'Sialkot': '51310', 'Narowal': '51600', 'Mardan': '23200', 'Attock': '43600', 'Sahiwal': '57000',
       'Burewala': '61010', 'Wah Cantt.': '47040', 'Daska': '51010', 'Kasur': '55050', 'Quetta': '87300',
       'Mianwali': '42200', 'Karachi': '74200', 'Muzaffarabad': '13100', 'Mirpur AJK': '10250', 'Wazirabad': '52000',
       'Vehari': '61100', 'Arifwala': '57450', 'Bahawalnagar': '62300', 'Lodhran': '59320', 'Khanewal': '58150',
       'Layyah': '31200', 'Jhang': '35200', 'Sheikhupura': '39350', 'Toba Tek Singh': '36050', 'Rahim Yar Khan': '64200',
       'Kharian': '50090', 'Dinga': '50280', 'Phalia': '50430', 'Taxila': '47080', 'Fateh Jang': '43350',
       'Haripur': '22620', 'Dina': '49400', 'Hyderabad': '71000', 'Sukkur': '65200', 'Murree': '47150',
       'Bahawalpur': '63100', 'Lalamusa': '50200', 'Chakwal': '48800', 'Wah': '47000', 'Joharabad': '41200',
       'hattar': '43394', 'Sarai Alamgir': '50000', 'Chiniot': '35400', 'Pir Mahal': '36300', 'Muridke': '39000',
       'Gujar Khan': '47850', 'Nowshera': '24100', 'Dara Adam Khel': '26100', 'Bannu': '28100', 'Mansehra': '21300',
       'Ali Pur Chatha': '52080', 'Mirpur Khas': '69000', 'Hazro': '43440'}

# surveyor e-mail aliases found in the baseline (form e-mail -> createdById used in the workbook)
EMAIL_ALIAS = {'shaperz.tauqeer@gmail.com': 'tauqeer@shapers.com.pk', 'shaperz.sw@gmail.com': 'waheed@shapers.com.pk'}

# floor name -> floor code (as in the baseline). Names not listed get no code (and an exception row).
FCODE = {'Ground Floor': 'G', 'First Floor': '1F', 'Second Floor': '2F', 'Basement': 'B1', 'Mezzanine Floor': 'M',
         'Lower Ground Floor': 'LG', 'Third Floor': '3F', 'Upper Ground Floor': 'UG', 'Forth Floor': '4F',
         'Second Floord': '2F', 'Basement 1': 'B1', 'Fourth Floor': '4F', 'Rooftop': 'RF', 'Basement 2': 'B2',
         'Eighth Floor': '8F', 'Third  Floor': '3F', 'Fifth Floor': '5F', 'Thirteenth Floor': '13F',
         'Sixth Floor': '6F', 'Seventh Floor': '7F', 'ground': 'G', 'Ninth Floor': '9F', 'Twelveth Floor': '12F',
         'Fourteenth Floor': '14F', '1st Floor': '1F', '1st floor': '1F', 'first floor': '1F', '2nd floor': '2F',
         '3rd floor': '3F', '4th floor': '4F', 'Roof top': 'RF', 'Tenth Floor': '10F', 'Eleventh Floor': '11F',
         'Fifteenth Floor': '15F', 'Sixteenth Floor': '16F', 'Seventeenth Floor': '17F', 'Eighteenth Floor': '18F',
         'Nineteenth Floor': '19F', 'Ground': 'G'}
DEFAULT_FLOOR_NAMES = ['Basement', 'Ground Floor', 'First Floor', 'Second Floor', 'Third Floor', 'Forth Floor',
                       'Fifth Floor', 'Sixth Floor', 'Seventh Floor', 'Eighth Floor', 'Ninth Floor', 'Tenth Floor',
                       'Eleventh Floor', 'Twelveth Floor', 'Thirteenth Floor', 'Fourteenth Floor',
                       'Fifteenth Floor', 'Sixteenth Floor', 'Seventeenth Floor', 'Eighteenth Floor',
                       'Nineteenth Floor']
ROLE_WORDS = {'owner', 'dealer', 'care taker', 'manager', 'caretaker', 'broker', "owner's representative",
              'representative', 'agent', 'contractor', 'friend', 'son', 'power of attorney', 'supervisor'}
MIME = {'.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.mp4': 'video/mp4'}

# land units accepted in field 138 (strict full-text match, like the baseline)
UW = {'kanal': 'Kanal', 'kanals': 'Kanal', 'marla': 'Marla', 'marlas': 'Marla', 'acre': 'Acre', 'acres': 'Acre',
      'sqft': 'Sq Ft', 'sq ft': 'Sq Ft', 'sq. ft': 'Sq Ft', 'sq.ft': 'Sq Ft', 'sqyds': 'Sq Yd', 'sq yd': 'Sq Yd',
      'sq. yd': 'Sq Yd', 'sq. yds': 'Sq Yd', 'sq yds': 'Sq Yd', 'sq yards': 'Sq Yd', 'sq. yards': 'Sq Yd',
      'sq.yards': 'Sq Yd', 'sq yard': 'Sq Yd', 'gazz': 'Sq Yd', 'gaz': 'Sq Yd', 'sqyd': 'Sq Yd', 'sq.yds': 'Sq Yd',
      'sq. yds.': 'Sq Yd'}
NUM = r'(\d+(?:,\d{3})*(?:\.\d+)?|\.\d+)'
SQ = r'(?:sq\.?\s*-?\s*ft\.?|sqft\.?|sft\.?|sq\.?\s*feet|square\s*feet)'

TK = ['26', '84', '85', '86', '87', '88', '89']      # floor-type slots (1..7)
SK = ['47', '48', '49', '50', '51']                  # floor size slots (1..5)
AK = ['4', '5', '6', '7', '8', '9', '10']            # floor area slots (1..7)


def first(meta, k):
    for kk, v in meta:
        if kk == k:
            return v
    return None


def allv(meta, k):
    return [v for kk, v in meta if kk == k]


def v_(meta, k):
    v = first(meta, k)
    if v is None:
        return None
    v = v.strip()
    return v or None


def junk_none(v):
    return None if (v is None or v in JUNK) else v


def num(s):
    f = float(s.replace(',', ''))
    return int(f) if f == int(f) else f


# ----------------------------------------------------------------------------------------------
# column rules (each verified against the baseline; see --revalidate)
# ----------------------------------------------------------------------------------------------
def r_area(meta):
    for k in ('114', '113', '115', '116', '117', '143'):
        v = v_(meta, k)
        if v and v not in JUNK:
            return v
    return junk_none(v_(meta, '124.3'))


def r_city(meta):
    c = v_(meta, '112')
    if c is None or c in JUNK:
        return None
    return CITYFIX.get(c, c)


PTYPE = {'plot': 'LAND_PLOT', 'office': 'OFFICE', 'house': 'RESIDENTIAL'}   # 'building' is left empty (as in baseline)


def dims_area(s):
    if not s:
        return None
    t = s.replace('×', 'x').replace('*', 'x')
    mm = re.fullmatch(r'\s*' + NUM + r'\s*(?:ft)?\s*[xX]\s*' + NUM + r'\s*(?:ft)?\s*', t)
    if not mm:
        return None
    r = round(num(mm.group(1)) * num(mm.group(2)), 4)
    return int(r) if r == int(r) else r


def unit_part(a):
    al = a.lower().strip()
    mm = re.match(r'\s*' + NUM + r'\s*(.*)$', al)
    if not mm:
        return None
    rest = re.sub(r'\(?\s*approx\.?\s*\)?', '', mm.group(2)).strip(' .#').strip()
    if rest in UW:
        return num(mm.group(1)), UW[rest]
    return None


def r_land(meta):
    a, d = v_(meta, '138'), v_(meta, '137')
    if a:
        r = unit_part(a)
        if r:
            return r
    x = dims_area(d)
    if x is not None:
        return x, 'Sq Ft'
    return None, None


def sqft_value(t):
    s = t.lower()
    s = re.sub(r'(?<=\d),(?=\d{3}(?!\d))', '', s)
    s = re.sub(r'approx(?:imately|imate)?|appprox|\(|\)', ' ', s)
    s = re.sub(r'\s+\.\s*', '.', s)
    s = re.sub(r'\s+', ' ', s).strip(' .,')
    mm = re.match(NUM + r'\s*(.*)$', s)
    if not mm:
        return None
    rest = mm.group(2).strip(' .,')
    if rest == '' or re.fullmatch(SQ + r'(?:\s*\.?)', rest) or rest in ('sft', 'sqft', 'sq ft', 'sq.ft', 'sq. ft'):
        return num(mm.group(1))
    return None


def legacy_gross(meta):
    t = v_(meta, '119')
    return sqft_value(t) if t else None


def parse_list(s):
    try:
        x = json.loads(s)
    except Exception:
        return None
    if isinstance(x, list):
        return [str(a).strip() for a in x if str(a).strip()]
    return None


def list141(meta):
    out = []
    for v in allv(meta, '141'):
        l = parse_list(v)
        if l:
            out = l
    return out


def wd(s):
    if not s:
        return None
    t = s.replace('×', 'x').replace('*', 'x')
    mm = re.fullmatch(r'\s*' + NUM + r'\s*(?:ft)?\s*[xX]\s*' + NUM + r'\s*(?:ft)?\s*', t)
    return (num(mm.group(1)), num(mm.group(2))) if mm else None


def areanum(s):
    if not s:
        return None
    mm = re.fullmatch(r'\s*' + NUM + r'\s*', s)
    if mm:
        return num(mm.group(1))
    return sqft_value(s)


def r_floors(meta):
    """list of floors: dict(name,w,d,area,note[,fb])"""
    mode = v_(meta, '139')
    N = areanum(v_(meta, '3'))
    L = list141(meta)
    types = [v_(meta, k) for k in TK]
    areas = [areanum(v_(meta, AK[k])) if v_(meta, AK[k]) else None for k in range(7)]
    na = sum(1 for k in AK if v_(meta, k))

    def mk(name, sz, ar):
        note, area = None, ar
        if sz:
            calc = round(sz[0] * sz[1], 4)
            calc = int(calc) if calc == int(calc) else calc
            if ar is not None and abs(ar - calc) > 1:
                note = 'legacy stated squareFeet=%s; using calculated floorWidthFt*floorDepthFt=%s' % (ar, calc)
            area = calc
        return dict(name=name, w=sz[0] if sz else None, d=sz[1] if sz else None, area=area, note=note,
                    stated=ar)

    out = []
    if mode == 'Yes':
        if any(types):
            n = N or sum(1 for t in types if t)
        elif N:
            n = N
        else:
            n = len(L)
        lastsz = lastar = None
        for k in range(int(n)):
            name = types[k] if k < 7 and types[k] else (
                L[k] if k < len(L) else (DEFAULT_FLOOR_NAMES[k] if 7 <= k < len(DEFAULT_FLOOR_NAMES) else None))
            sz = wd(v_(meta, SK[k])) if k < 5 else None
            ar = areas[k] if k < 7 else None
            if k >= 5:
                if sz:
                    lastsz = sz
                else:
                    sz = lastsz
            else:
                lastsz = sz or lastsz
            if k >= 7:
                if ar is not None:
                    lastar = ar
                else:
                    ar = lastar
            else:
                lastar = ar if ar is not None else lastar
            out.append(mk(name, sz, ar))
        return out
    if na == 0:
        return [mk(nm, None, None) for nm in L]
    firstar = next((a for a in areas if a is not None), None)
    cnt = int(N or len(L) or 1)
    if na < cnt:
        r = [mk(None, None, firstar)]
        r[0]['fb'] = True
        return r
    names = L if len(L) == cnt else [None] * cnt
    for k in range(cnt):
        ar = areas[k] if k < 7 else None
        sz = None
        if cnt == 1 and names[0] != 'Plot':
            cs = wd(v_(meta, '142'))
            if cs and ar is not None and round(cs[0] * cs[1], 4) == ar:
                sz = cs
        out.append(mk(names[k], sz, ar))
    return out


def r_gross(meta, fl=None):
    fl = r_floors(meta) if fl is None else fl
    cf = [f for f in fl if FCODE.get(f['name'])] if fl and not fl[0].get('fb') else []
    if cf and all(f['area'] is not None for f in cf):
        return sum(f['area'] for f in cf)
    return legacy_gross(meta)


def r_nfloors(meta, fl=None):
    N = areanum(v_(meta, '3'))
    if N is None:
        return None
    fl = r_floors(meta) if fl is None else fl
    if not fl or fl[0].get('fb'):
        return N
    ex = sum(1 for f in fl if f['name'] and f['name'] != 'Plot' and not FCODE.get(f['name']))
    return N - ex if N - ex > 0 else N


LOC0 = re.compile(r'^\s*(near|nearby|opp|opposite|beside|besides|adjacent|along|next|behind|infront|facing)\b', re.I)
LOC1 = re.compile(r'\b(near|opp|opposite|beside|besides|beaside)\b', re.I)


def _clean_brand(x):
    return re.sub(r'\s*\betc\b\.?$', '', x.strip(), flags=re.I).strip().rstrip('.').strip()


def _brand_piece(t):
    t = t.strip()
    core = t.rstrip(' ,')
    if ',' not in core:
        c = _clean_brand(core)
        return ([c] if c else []), False
    parts = [x.strip() for x in core.split(',')]
    if LOC0.search(core) or LOC1.search(core) or ',,' in core or any(
            len(p.split(' ')[0]) == 1 for p in parts if p):
        return [t], True            # ambiguous free text: kept whole
    return [c for c in (_clean_brand(x) for x in parts) if c], False


def r_neighbors(meta):
    items, ambiguous = [], False
    v = v_(meta, '125')
    if v:
        l = parse_list(v)
        if l:
            items += l
    for k in ('118', '120'):
        t = v_(meta, k)
        if t:
            p, amb = _brand_piece(t)
            items += p
            ambiguous = ambiguous or amb
    seen = []
    for x in items:
        if x not in seen:
            seen.append(x)
    return (json.dumps(seen, ensure_ascii=False) if seen else None), ambiguous


def r_amount(meta, k):
    t = v_(meta, k)
    if t and re.fullmatch(r'\d[\d,]*(\.\d+)?', t):
        return num(t)
    return None


def norm_phone(p):
    if not p:
        return ''
    d = re.sub(r'\D', '', p)
    if d.startswith('92') and len(d) >= 11:
        return '+' + d
    if d.startswith('0') and 10 <= len(d) <= 11:
        return '+92' + d[1:]
    if len(d) == 10 and d.startswith('3'):
        return '+92' + d
    return ''


def parse_dt(s):
    if not s or s == 'NULL' or s.startswith('0000'):
        return None
    try:
        return datetime.datetime.strptime(s[:19], '%Y-%m-%d %H:%M:%S')
    except ValueError:
        return None


def r_media(meta):
    """all media URLs found (images field 83, videos field 103) in order"""
    out = []
    for key, typ in (('83', 'IMAGE'), ('103', 'VIDEO')):
        for v in allv(meta, key):
            v = (v or '').strip()
            if not v:
                continue
            try:
                l = json.loads(v)
            except Exception:
                l = None
            if not isinstance(l, list):
                l = [v]
            for u in l:
                if isinstance(u, str) and u.strip():
                    u = u.strip()
                    fn = urllib.parse.unquote(os.path.basename(urllib.parse.urlparse(u).path))
                    ext = os.path.splitext(fn)[1].lower()
                    out.append(dict(type=typ, url=u, filename=fn, mime=MIME.get(ext), ext=ext))
    return out


def baseline_media_rule(media_list):
    """baseline keeps the first IMAGE and the first VIDEO that has a known file extension"""
    keep = []
    for t in ('IMAGE', 'VIDEO'):
        for x in media_list:
            if x['type'] == t and x['mime']:
                keep.append(x)
                break
    return keep


# --all-media: set in main() before anything is built. Keeps EVERY usable image/video of a survey instead of the first
# image + first video only.
ALL_MEDIA = False
ALL_MEDIA_REPORT = {}
IMAGE_LIMIT_BYTES = 10 * 1024 * 1024     # the app rejects larger images / videos on upload
VIDEO_LIMIT_BYTES = 300 * 1024 * 1024
BROWSER_UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
              'Chrome/124.0 Safari/537.36')           # the portal host answers 406 to short user agents


def media_key(url):
    """identity of a file for de-duplication / matching: no scheme, percent-decoding undone"""
    return urllib.parse.unquote(re.sub(r'^https?://', '', url.strip(), flags=re.I))


def r_media_all(meta):
    """every usable image/video of an entry: field 83 (pictures), 103 (videos), then 129 (backup documents /
    original pictures & videos). Only URLs with a known image/video extension are kept (PDFs and placeholder links
    are not media); a file listed twice counts once. The type comes from the extension."""
    out, seen = [], set()
    for key in ('83', '103', '129'):
        for v in allv(meta, key):
            v = (v or '').strip()
            if not v:
                continue
            try:
                l = json.loads(v)
            except Exception:
                l = None
            if not isinstance(l, list):
                l = [v]
            for u in l:
                if not (isinstance(u, str) and u.strip()):
                    continue
                u = u.strip()
                fn = urllib.parse.unquote(os.path.basename(urllib.parse.urlparse(u).path))
                ext = os.path.splitext(fn)[1].lower()
                mime = MIME.get(ext)
                if not mime or media_key(u) in seen:
                    continue
                seen.add(media_key(u))
                out.append(dict(type='IMAGE' if mime.startswith('image/') else 'VIDEO', url=u, filename=fn,
                                mime=mime, ext=ext, field=key))
    return out


# ----------------------------------------------------------------------------------------------
# build all sheet rows for one entry
# ----------------------------------------------------------------------------------------------
def build(entry_id, meta, entry, known_surveyors):
    """returns dict(survey=dict, floors=[dict], contact=dict, attachments=[dict], exceptions=[tuple], flags=[str])"""
    exc = []
    flags = []

    def E(entity, field, raw, norm, reason, sev, note=None):
        exc.append((entity, entry_id, field, raw, norm, reason, sev, note))

    s = {}
    fl = r_floors(meta)
    s['legacySurveyId'] = entry_id
    s['address'] = junk_none(v_(meta, '111'))
    # latitude and longitude are two separate numeric columns in the workbook (never JSON): 124.2 = latitude, 124.1 = longitude
    try:
        s['latitude'] = round(float(v_(meta, '124.2')), 6)
        s['longitude'] = round(float(v_(meta, '124.1')), 6)
    except (TypeError, ValueError):
        s['latitude'] = s['longitude'] = None
    s['city'] = r_city(meta)
    s['area'] = r_area(meta)
    s['zipCode'] = ZIP.get(s['city']) if s['city'] else None
    s['propertyType'] = PTYPE.get(v_(meta, '136'))
    s['ownerName'] = junk_none(v_(meta, '68'))
    la, lu = r_land(meta)
    s['landArea'], s['landAreaUnit'] = la, lu
    s['grossBuildingArea'] = r_gross(meta, fl)
    s['numberOfFloors'] = r_nfloors(meta, fl)
    s['availabilityStatus'] = 'active'
    s['propertyStatus'] = {'Yes': 'VACANT', 'No': 'RENTED'}.get(v_(meta, '133'))
    s['availableFor'] = junk_none(v_(meta, '134'))
    s['advanceAmount'] = r_amount(meta, '75')
    s['securityDepositAmount'] = r_amount(meta, '77')
    s['possessionTerms'] = junk_none(v_(meta, '78'))
    dc = parse_dt(entry.get('date_created'))
    s['surveyDate'] = dc.replace(hour=0, minute=0, second=0) if dc else None
    s['createdAt'] = dc
    nb, nb_amb = r_neighbors(meta)
    s['neighboringBusinesses'] = nb
    em = v_(meta, '110')
    cb = None
    if em and em != 'Select':
        cb = EMAIL_ALIAS.get(em, em)
    s['createdById'] = cb
    s['freshSurveyCount'] = 0

    # ---- survey-level exceptions (same codes/wording as the baseline's 05_Seed_Exceptions)
    if dc is None:
        E('Survey', 'createdAt', entry.get('date_created'), None, 'INVALID_LEGACY_DATE', 'HIGH',
          'No valid date_created; createdAt is blank in the seed')
    if s['latitude'] is None:
        E('Survey', 'latitude/longitude', '%s,%s' % (v_(meta, '124.2') or '', v_(meta, '124.1') or ''), None,
          'MISSING_COORDINATES', 'MEDIUM')
    else:
        la_, lo_ = s['latitude'], s['longitude']
        if not (23.5 <= la_ <= 37.2 and 60.5 <= lo_ <= 77.9):
            E('Survey', 'latitude/longitude', '%s,%s' % (v_(meta, '124.2'), v_(meta, '124.1')),
              '%s,%s' % (la_, lo_), 'COORDINATES_OUTSIDE_PAKISTAN', 'MEDIUM', 'Kept; verify')
    if cb is None:
        E('Survey', 'createdById', em, None, 'BUSINESS_DEVELOPER_MISSING', 'MEDIUM',
          'No Business_Developer email on this record; createdById left blank')
    elif cb not in known_surveyors:
        E('Survey', 'createdById', em, cb, 'NEW_SURVEYOR_EMAIL', 'MEDIUM',
          'Surveyor e-mail not present in the baseline workbook; value kept as-is, CRE user lookup needed '
          '(no user is created by this script)')
    if s['city'] and s['zipCode'] is None:
        E('Survey', 'zipCode', s['city'], None, 'NEW_CITY_NO_ZIP', 'LOW',
          'City not in the baseline city->zip list; zipCode left blank')
    for k, label in (('75', 'advance'), ('77', 'security')):
        t = v_(meta, k)
        if t and r_amount(meta, k) is None:
            mm = re.fullmatch(r'(\d+)\s*months?', t, flags=re.I)
            E('Survey', label, t, None, 'COMMERCIAL_AMOUNT_UNPARSEABLE', 'LOW',
              '%s months of rent, not a currency amount' % mm.group(1) if mm else 'unparseable as a currency amount')
    if v_(meta, '133') == 'No' and v_(meta, '78') and 'available' in v_(meta, '78').lower():
        E('Survey', 'propertyStatus', 'No | %s' % v_(meta, '78'), 'RENTED', 'CONFLICT_WITH_POSSESSION', 'HIGH',
          'propertyStatus kept per approved Available_Not_Rented mapping')
    coded = [f for f in fl if FCODE.get(f['name'])] if fl and not fl[0].get('fb') else []
    lg = legacy_gross(meta)
    if coded:
        measured = [f for f in coded if f['area'] is not None]
        if len(measured) == len(coded):
            if lg is not None and lg != s['grossBuildingArea']:
                E('Survey', 'grossBuildingArea', lg, s['grossBuildingArea'], 'GROSS_AREA_MISMATCH', 'MEDIUM',
                  'Recalculated from %d SurveyFloor rows; legacy value replaced' % len(coded))
        else:
            E('Survey', 'grossBuildingArea', lg, None, 'FLOOR_TOTAL_INCOMPLETE', 'LOW',
              '%d/%d floors have a measured area (partial sum %s); legacy grossBuildingArea/numberOfFloors kept as-is'
              % (len(measured), len(coded), sum(f['area'] for f in measured)))
    elif fl and fl[0].get('fb'):
        pass
    # ---- floors
    frows = []
    for k, f in enumerate(fl):
        frows.append(dict(legacySurveyId=entry_id, floorCode=FCODE.get(f['name']), floorName=f['name'], sortOrder=k,
                          floorWidthFt=f['w'], floorDepthFt=f['d'], floorAreaSqFt=f['area'], notes=f['note']))
        if f['name'] and not FCODE.get(f['name']):
            E('SurveyFloor', 'floorName', f['name'], None, 'FLOOR_NAME_UNMAPPED', 'LOW',
              'No confident floor-code mapping; kept as floorName only')
        if f['note']:
            E('SurveyFloor', 'floorAreaSqFt', f['stated'], f['area'], 'FLOOR_AREA_MISMATCH', 'MEDIUM', f['note'])
    # ---- contact
    nm = junk_none(v_(meta, '68'))
    contact = dict(legacySurveyId=entry_id, name=nm,
                   role=(lambda r: None if (r is None or (r in JUNK and r != 'nil')) else r)(v_(meta, '69')),
                   phone=junk_none(v_(meta, '70')), email=None, createdAt=dc)
    cp = contact['role']
    raw_contact = json.dumps({'Owner_Name': v_(meta, '68'), 'Contact_Person': v_(meta, '69'),
                              'Contact_Number': v_(meta, '70')}, ensure_ascii=False, separators=(',', ':'))
    ph = v_(meta, '70')
    strict = re.findall(r'(?<!\d)(?:\+?92|0)?3\d{2}[\s\-]?\d{7}(?!\d)', ph) if ph else []
    HOLD = 'Previously HOLD: %s. Row kept in 02_Survey_Contacts with raw values'
    bad_name = bool(nm) and (re.fullmatch(r'(mr|mrs|ms|dr|miss)?[\s\.\-]*', nm, flags=re.I) is not None
                             or not re.search(r'[A-Za-z]{2}', nm))
    if cp and cp.strip().lower() not in ROLE_WORDS:
        E('SurveyContact', 'contact', raw_contact,
          json.dumps({'name': '', 'role': '', 'phone': norm_phone(ph)}, ensure_ascii=False, separators=(',', ':')),
          'REVIEW_CP_SEMANTICS', 'MEDIUM',
          HOLD % 'cp is not a pure role word (may be an internal BD name)')
    elif len(strict) > 1:
        E('SurveyContact', 'contact', raw_contact,
          json.dumps({'name': '', 'role': cp or '', 'phone': ''}, ensure_ascii=False, separators=(',', ':')),
          'REVIEW_MULTI_PHONE', 'MEDIUM', HOLD % 'Multiple strict phones in one cell; do not split')
    elif bad_name:
        E('SurveyContact', 'contact', raw_contact,
          json.dumps({'name': '', 'role': cp or '', 'phone': norm_phone(ph)},
                     ensure_ascii=False, separators=(',', ':')), 'INVALID_NAME', 'MEDIUM',
          HOLD % 'Required name is title/placeholder; do not fabricate')
    # ---- media
    allm = r_media(meta)
    keep = r_media_all(meta) if ALL_MEDIA else baseline_media_rule(allm)
    arows = [dict(legacySurveyId=entry_id, attachmentType=x['type'], originalUrl=x['url'],
                  originalFilename=x['filename'], mimeType=x['mime'], fileSizeBytes=None) for x in keep]
    if not any(x['type'] == 'IMAGE' for x in keep):
        E('Attachment', 'IMAGE', None, None, 'NEW_MEDIA_NO_IMAGE', 'MEDIUM', 'No usable image URL on this entry')
    for x in allm:
        if not x['mime']:
            E('Attachment', x['type'], x['url'], None, 'NEW_MEDIA_URL_NO_KNOWN_EXTENSION', 'LOW',
              'URL has no .jpg/.jpeg/.png/.mp4 extension (placeholder or unsupported); not listed in 04_Attachments')
    extras = [] if ALL_MEDIA else [x for x in allm if x['mime'] and x not in keep]
    if extras:
        E('Attachment', 'additional media', ' | '.join(x['url'] for x in extras), len(extras),
          'NEW_MEDIA_EXTRA_NOT_LISTED', 'LOW',
          'Baseline rule lists only the first image and first video per survey; these further URLs are recorded here '
          'only (not downloaded)')
    docs = [u for v in allv(meta, '129') for u in (parse_list(v) or [])
            if os.path.splitext(urllib.parse.urlparse(u).path)[1]
            and not (ALL_MEDIA and os.path.splitext(urllib.parse.urlparse(u).path)[1].lower() in MIME)]
    if docs:
        E('Attachment', 'backup documents (field 129)', ' | '.join(docs), len(docs), 'NEW_MEDIA_BACKUP_DOCUMENT',
          'LOW', 'Backup-document uploads have no attachment type in the workbook; URLs recorded only')
    # ---- data-quality notes that need a human look
    if nb_amb:
        E('Survey', 'neighboringBusinesses', ' | '.join(filter(None, [v_(meta, '118'), v_(meta, '120')])), nb,
          'NEW_NEIGHBOR_TEXT_AMBIGUOUS', 'LOW',
          'Free-text brands contain commas/location words; kept as a single entry (baseline-style heuristic). Review.')
    dup = [k for k in ('68', '69', '70', '111', '112', '119', '3', '139', '134', '133', '110')
           if len(set(a for a in allv(meta, k) if a and a.strip())) > 1]
    if dup:
        E('Survey', 'duplicate meta keys', ','.join(dup), None, 'NEW_DUPLICATE_META_KEYS', 'MEDIUM',
          'Field saved more than once with different values; first non-empty value used')
    ps = v_(meta, '137')
    if ps and dims_area(ps) is not None:
        E('Survey', 'plotSize', ps, dims_area(ps), 'LAND_DIMENSION_ORDER_UNKNOWN', 'LOW',
          'Legacy plotSize does not establish which number is frontage vs. depth; area is safe to compute '
          '(commutative) but landFrontageFt/landDepthFt are left null')
    elif ps and s['landArea'] is None:
        E('Survey', 'plotSize', ps, None, 'PLOT_DIMENSIONS_UNPARSEABLE', 'LOW',
          'Neither a recognizable dimension pair nor a known area-unit expression')
    cs = v_(meta, '142')
    if cs and v_(meta, '139') != 'Yes' and wd(cs) is not None:
        w_, d_ = wd(cs)
        carea = w_ * d_
        carea = int(carea) if carea == int(carea) else carea
        phys = len([f for f in fl if FCODE.get(f['name'])])
        if phys > 1:
            E('SurveyFloor', 'coveredSize', cs, None, 'COVERED_SIZE_MULTI_FLOOR_AMBIGUOUS', 'LOW',
              '%d physical floor(s); relationship between coveredSize and a specific floor is not deterministic' % phys)
        elif phys == 1 and len(fl) == 1 and fl[0]['name'] != 'Plot':
            st = fl[0]['stated']
            if st is not None and st == carea:
                E('SurveyFloor', 'coveredSize', cs, carea, 'COVERED_SIZE_RECOVERED_FLOOR_DIMENSIONS', 'LOW',
                  'Single physical floor; coveredSize area matches floorAreaSqFt=%s' % st)
            else:
                E('SurveyFloor', 'coveredSize', cs, carea, 'COVERED_SIZE_AREA_MISMATCH', 'LOW',
                  'coveredSize area %s does not match floorAreaSqFt=%s; dimensions left null'
                  % (carea, ('%.1f' % st) if isinstance(st, (int, float)) else None))
    return dict(survey=s, floors=frows, contact=contact, attachments=arows, exceptions=exc)


# ----------------------------------------------------------------------------------------------
# sources
# ----------------------------------------------------------------------------------------------
def load_csv(entry_csv, meta_csv, only_ids=None):
    entries, metas = {}, collections.defaultdict(list)
    with open(entry_csv, newline='', encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            if str(row.get('form_id', FORM_ID)) != str(FORM_ID):
                continue
            i = int(row['id'])
            if only_ids is None or i in only_ids:
                entries[i] = row
    with open(meta_csv, newline='', encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            i = int(row['entry_id'])
            if i in entries:
                v = row['meta_value']
                metas[i].append((row['meta_key'], None if v == 'NULL' else v))
    return entries, metas


def load_dirs(dirs):
    """scan folders (recursively) for CSV exports; recognise entry / meta files by their header row,
    merge several files and drop duplicate rows (overlapping dumps are fine)."""
    entries, meta_rows = {}, {}
    seen_files = collections.Counter()
    for d in dirs:
        for root, _, files in os.walk(d):
            for fn in sorted(files):
                if not fn.lower().endswith('.csv'):
                    continue
                path = os.path.join(root, fn)
                with open(path, newline='', encoding='utf-8-sig') as f:
                    rd = csv.DictReader(f)
                    cols = set(rd.fieldnames or [])
                    if {'entry_id', 'meta_key', 'meta_value'} <= cols:
                        kind = 'meta'
                        for row in rd:
                            try:
                                rid = int(row.get('id') or 0)
                            except ValueError:
                                continue
                            if str(row.get('form_id', FORM_ID)) == str(FORM_ID):
                                meta_rows[rid or (int(row['entry_id']), row['meta_key'], row['meta_value'])] = row
                    elif {'id', 'form_id', 'date_created'} <= cols:
                        kind = 'entry'
                        for row in rd:
                            if str(row['form_id']) == str(FORM_ID):
                                entries[int(row['id'])] = row
                    else:
                        continue
                    seen_files[kind] += 1
                    print('  read %-5s file: %s' % (kind, os.path.relpath(path, d)))
    if not entries or not meta_rows:
        sys.exit('No entry/meta CSV files found in %s  (found: %s)' % (dirs, dict(seen_files)))
    metas = collections.defaultdict(list)
    for rid in sorted(meta_rows, key=lambda k: (k if isinstance(k, int) else 0)):
        row = meta_rows[rid]
        i = int(row['entry_id'])
        if i in entries:
            v = row['meta_value']
            metas[i].append((row['meta_key'], None if v == 'NULL' else v))
    return entries, metas


def load_mysql(prefix, ids_filter_sql, params=()):
    try:
        import pymysql as driver
        connect = lambda **kw: driver.connect(**kw)
    except ImportError:
        try:
            import mysql.connector as driver
            connect = lambda **kw: driver.connect(**kw)
        except ImportError:
            sys.exit("A MySQL driver is needed for --mysql:  pip install pymysql")
    host = os.environ.get('GF_DB_HOST', 'localhost')
    name = os.environ.get('GF_DB_NAME')
    user = os.environ.get('GF_DB_USER')
    pw = os.environ.get('GF_DB_PASS') or getpass.getpass('Database password (not stored): ')
    if not name or not user:
        sys.exit("Set GF_DB_NAME and GF_DB_USER (and optionally GF_DB_HOST / GF_DB_PASS) environment variables.")
    kw = dict(host=host, user=user, password=pw, port=int(os.environ.get('GF_DB_PORT', '3306')))
    kw['database' if 'mysql.connector' in driver.__name__ else 'db'] = name
    cn = connect(**kw)
    cur = cn.cursor()
    try:                                   # belt and braces: make the session read-only
        cur.execute("SET SESSION TRANSACTION READ ONLY")
    except Exception:
        pass
    cur.execute("SELECT id FROM %sgf_entry WHERE form_id=%%s" % prefix, (FORM_ID,))
    all_ids = {int(r[0]) for r in cur.fetchall()}
    return cn, cur, all_ids


def fetch_mysql_entries(cur, prefix, ids):
    entries, metas = {}, collections.defaultdict(list)
    ids = sorted(ids)
    for a in range(0, len(ids), 500):
        chunk = ids[a:a + 500]
        ph = ','.join(['%s'] * len(chunk))
        cur.execute("SELECT id,form_id,date_created,status,created_by FROM %sgf_entry WHERE id IN (%s)" % (prefix, ph),
                    chunk)
        for r in cur.fetchall():
            dcv = r[2]
            entries[int(r[0])] = dict(id=r[0], form_id=r[1], status=r[3], created_by=r[4],
                                      date_created=dcv.strftime('%Y-%m-%d %H:%M:%S') if hasattr(dcv, 'strftime')
                                      else (str(dcv) if dcv is not None else None))
        cur.execute("SELECT entry_id,meta_key,meta_value FROM %sgf_entry_meta WHERE entry_id IN (%s) ORDER BY id" %
                    (prefix, ph), chunk)
        for r in cur.fetchall():
            metas[int(r[0])].append((str(r[1]), r[2]))
    return entries, metas


# ----------------------------------------------------------------------------------------------
# workbook I/O
# ----------------------------------------------------------------------------------------------
DATE_FMT = 'yyyy\\-mm\\-dd\\ hh:mm:ss'


def read_baseline(path):
    wb = openpyxl.load_workbook(path)
    ws = wb['01_Surveys']
    hdr = [c.value for c in ws[1]]
    ids = set()
    surveyors = set()
    ci = hdr.index('createdById')
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            ids.add(int(row[0]))
            if row[ci]:
                surveyors.add(row[ci])
    return wb, ids, surveyors


def append_rows(ws, rows):
    hdr = [c.value for c in ws[1]]
    style_row = ws[ws.max_row]
    for r in rows:
        ws.append([r.get(h) for h in hdr])
        for c in ws[ws.max_row]:
            tmpl = style_row[c.column - 1]
            if tmpl.has_style:
                c._style = tmpl._style
            if isinstance(c.value, datetime.datetime):
                c.number_format = DATE_FMT


EXC_HDR = ['entity', 'legacySurveyId', 'field', 'rawValue', 'normalizedValue', 'reason', 'severity', 'notes']


# ----------------------------------------------------------------------------------------------
# --all-media: rebuild 04_Attachments for EVERY survey (not only the new ones)
# ----------------------------------------------------------------------------------------------
MEDIA_EXC_REASONS = {'NEW_MEDIA_EXTRA_NOT_LISTED', 'NEW_MEDIA_BACKUP_DOCUMENT'}


def plan_all_media(wb, entries, metas, base_ids):
    """For every survey already in the workbook that is still an ACTIVE entry in the source: its full media list.
    A file already in 04_Attachments is reused untouched (same URL text, so a file already downloaded keeps its local
    name); any other file is new (size empty until --check-urls fills it); an old row whose file is no longer on the
    entry is dropped (replaced or deleted on the portal). Surveys that are missing from the source or not active keep
    their rows exactly as they are.
    returns (rows, changes, stats, untouched_ids)"""
    ws = wb['04_Attachments']
    hdr = [c.value for c in ws[1]]
    old_by = collections.defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is not None:
            old_by[int(r[0])].append(dict(zip(hdr, r)))
    rows, changes, untouched = [], [], set()
    stats = collections.Counter()
    for sid in sorted(base_ids):
        e = entries.get(sid)
        if e is None or (e.get('status') or 'active') != 'active' or sid not in metas:
            untouched.add(sid)
            rows += old_by.get(sid, [])
            stats['surveys_kept_as_is'] += 1
            continue
        have = {media_key(r['originalUrl']): r for r in old_by.get(sid, [])}
        wanted = set()
        for x in r_media_all(metas[sid]):
            k = media_key(x['url'])
            wanted.add(k)
            if k in have:
                rows.append(have[k])
                stats['unchanged'] += 1
            else:
                rows.append(dict(legacySurveyId=sid, attachmentType=x['type'], originalUrl=x['url'],
                                 originalFilename=x['filename'], mimeType=x['mime'], fileSizeBytes=None))
                stats['added'] += 1
                changes.append(('ADDED', sid, x['type'], '', x['url'], 'field %s' % x['field']))
        for k, r in have.items():
            if k not in wanted:
                stats['removed'] += 1
                changes.append(('REMOVED', sid, r['attachmentType'], r['originalUrl'], '',
                                'no longer on the entry (replaced or deleted on the portal)'))
    return rows, changes, stats, untouched


def check_media_rows(rows, workers=8):
    """HEAD every row that has no size yet (= new files), 8 at a time, no cookies. Fills fileSizeBytes. A file counts as
    good only with HTTP 200 and the right content type; a redirect to wp-login.php or a 404 is a failure.
    returns [(row, reason)] for the failures"""
    import concurrent.futures, time, urllib.error, urllib.request
    todo = [r for r in rows if not r.get('fileSizeBytes')]

    def one(r):
        url = r['originalUrl']
        want = str(r['attachmentType']).lower() + '/'
        last = 'unknown error'
        for attempt in range(3):
            try:
                try:
                    resp = urllib.request.urlopen(urllib.request.Request(
                        url, method='HEAD', headers={'User-Agent': BROWSER_UA, 'Accept': '*/*'}), timeout=30)
                except urllib.error.HTTPError as he:
                    if he.code not in (403, 405, 406, 501):
                        raise
                    resp = urllib.request.urlopen(urllib.request.Request(
                        url, headers={'User-Agent': BROWSER_UA, 'Range': 'bytes=0-0'}), timeout=30)
                with resp:
                    if 'wp-login.php' in resp.geturl():
                        return 'redirects to the WordPress login (file missing or protected)'
                    ct = (resp.headers.get('Content-Type') or '').split(';')[0].strip().lower()
                    if not ct.startswith(want):
                        return 'content-type "%s", expected %s*' % (ct, want)
                    cr = resp.headers.get('Content-Range')
                    cl = resp.headers.get('Content-Length')
                    if cr and '/' in cr and cr.rsplit('/', 1)[1].isdigit():
                        size = int(cr.rsplit('/', 1)[1])
                    else:
                        size = int(cl) if cl and cl.isdigit() else 0
                    if not size:
                        return 'the server reported no file size'
                    r['fileSizeBytes'] = size
                    return None
            except urllib.error.HTTPError as he:
                last = 'HTTP %d' % he.code
                if he.code in (404, 410):
                    return last
            except Exception as ex:
                last = '%s: %s' % (type(ex).__name__, str(ex)[:80])
            time.sleep(2 * (attempt + 1))
        return last

    failures = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as ex:
        for n, (r, reason) in enumerate(zip(todo, ex.map(one, todo)), 1):
            if reason:
                failures.append((r, reason))
            if n % 500 == 0:
                print('  checked %d / %d new files' % (n, len(todo)))
    print('Checked %d new media URLs (headers only, nothing downloaded): %d ok, %d failed'
          % (len(todo), len(todo) - len(failures), len(failures)))
    return failures


def replace_rows(ws, rows):
    """replace every data row of a sheet (header and column styles kept)"""
    hdr = [c.value for c in ws[1]]
    styles = [copy.copy(c._style) for c in ws[2]] if ws.max_row >= 2 else None
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row)
    for n, r in enumerate(rows, start=2):
        for j, h in enumerate(hdr, start=1):
            c = ws.cell(row=n, column=j, value=r.get(h))
            if styles:
                c._style = copy.copy(styles[j - 1])
    ws._current_row = ws.max_row


def refresh_media_exceptions(ws, source_ids, entries, metas):
    """05_Seed_Exceptions: the two media notes written by the first-image/first-video rule ('further URLs not listed',
    'backup documents') are out of date once all media is listed, so they are removed for every survey that is rebuilt,
    and the backup-document note is written again for the files that are NOT images/videos (PDFs and the like).
    returns (rows to append, number of old rows removed)"""
    hdr = [c.value for c in ws[1]]
    ri, ii = hdr.index('reason'), hdr.index('legacySurveyId')
    drop = [n for n, r in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1)
            if n > 1 and r[ri] in MEDIA_EXC_REASONS and r[ii] is not None and int(r[ii]) in source_ids]
    for n in reversed(drop):
        ws.delete_rows(n)
    ws._current_row = ws.max_row
    fresh = []
    for sid in sorted(source_ids):
        docs = [u for v in allv(metas.get(sid, []), '129') for u in (parse_list(v) or [])
                if os.path.splitext(urllib.parse.urlparse(u).path)[1]
                and os.path.splitext(urllib.parse.urlparse(u).path)[1].lower() not in MIME]
        if docs:
            fresh.append(dict(zip(EXC_HDR, ('Attachment', sid, 'backup documents (field 129)', ' | '.join(docs),
                                            len(docs), 'NEW_MEDIA_BACKUP_DOCUMENT', 'LOW',
                                            'Backup uploads that are not images/videos (e.g. PDF); URLs recorded only'))))
    return fresh, len(drop)


# ----------------------------------------------------------------------------------------------
# Google Sheets (append-only)
# ----------------------------------------------------------------------------------------------
def open_gsheet(sheet_id, creds):
    try:
        import gspread
    except ImportError:
        sys.exit("pip install gspread   (needed for --gsheet-id)")
    if not creds or not os.path.exists(creds):
        sys.exit("--gsheet-creds (or env GSHEET_CREDS) must point to the service-account JSON key file")
    return gspread.service_account(filename=creds).open_by_key(sheet_id)


def read_gsheet(sh):
    ws = sh.worksheet('01_Surveys')
    hdr = ws.row_values(1)
    ids = {int(float(x)) for x in ws.col_values(1)[1:] if str(x).strip()}
    ci = hdr.index('createdById') + 1
    surveyors = {x for x in ws.col_values(ci)[1:] if x}
    return ids, surveyors


def _serial(dt):
    return (dt - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0


def gsheet_append(sh, sheets):
    """append only; existing rows are never touched. Re-checks the last ID right before writing."""
    for name, rows in sheets.items():
        if not rows:
            continue
        ws = sh.worksheet(name)
        hdr = ws.row_values(1)
        vals, datecols = [], set()
        for r in rows:
            line = []
            for j, h in enumerate(hdr):
                v = r.get(h)
                if isinstance(v, datetime.datetime):
                    v = _serial(v); datecols.add(j)
                line.append('' if v is None else v)
            vals.append(line)
        start = len(ws.col_values(1)) + 1
        ws.append_rows(vals, value_input_option='RAW', table_range='A1')
        end = start + len(vals) - 1
        for j in datecols:
            col = chr(ord('A') + j) if j < 26 else 'A' + chr(ord('A') + j - 26)
            ws.format('%s%d:%s%d' % (col, start, col, end),
                      {'numberFormat': {'type': 'DATE_TIME', 'pattern': 'yyyy-mm-dd hh:mm:ss'}})
        print('  appended %d rows to %s' % (len(vals), name))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--baseline', help='existing historical-surveys.xlsx (read only)')
    ap.add_argument('--gsheet-id', help='Google Sheet ID: append new rows straight into it (no xlsx)')
    ap.add_argument('--gsheet-creds', default=os.environ.get('GSHEET_CREDS'), help='service-account JSON key file')
    ap.add_argument('--out', help='NEW workbook to write (must differ from --baseline)')
    ap.add_argument('--in-place', action='store_true',
                    help='append the new rows INTO the --baseline file itself (a dated backup copy is made first)')
    ap.add_argument('--entry-csv'); ap.add_argument('--meta-csv')
    ap.add_argument('--dump-dir', action='append',
                    help='folder with the exported CSV files (searched recursively; repeat the option for several '
                         'folders). Entry/meta files are recognised automatically and merged.')
    ap.add_argument('--mysql', action='store_true'); ap.add_argument('--prefix', default='wp_')
    ap.add_argument('--dry-run', action='store_true', help='pre-flight report only')
    ap.add_argument('--only-new-workbook', help='additionally write a workbook with ONLY the new rows')
    ap.add_argument('--check-urls', action='store_true',
                    help='HTTP HEAD each listed media URL (headers only) to fill fileSizeBytes / flag broken links')
    ap.add_argument('--revalidate', action='store_true',
                    help='if the source also contains entries already in the baseline, re-derive them and report '
                         'how many cells differ (accuracy check of the rules)')
    ap.add_argument('--summary', default=None, help='summary text file (default: next to --out)')
    ap.add_argument('--all-media', action='store_true',
                    help='keep EVERY image and video of every survey (fields 83, 103 and 129), not only the first image '
                         'and first video, and rebuild 04_Attachments for the surveys already in the baseline too '
                         '(replaced / deleted / newly added files). Existing rows whose file is unchanged are kept as '
                         'they are. Needs --baseline. Use with --check-urls so new files get a size and dead links are '
                         'dropped.')
    a = ap.parse_args()
    global ALL_MEDIA
    ALL_MEDIA = a.all_media
    if a.all_media and (a.gsheet_id or not a.baseline):
        ap.error('--all-media works on an xlsx workbook: give --baseline (and not --gsheet-id)')
    if a.in_place:
        if not a.baseline or a.out:
            ap.error('--in-place needs --baseline and no --out')
        a.out = os.path.splitext(a.baseline)[0] + '.__tmp__.xlsx'
    if not a.baseline and not a.gsheet_id:
        ap.error('give --baseline (xlsx) or --gsheet-id')
    if not a.dry_run and not a.out and not a.gsheet_id:
        ap.error('--out is required unless --dry-run or --gsheet-id')
    if a.out and a.baseline and not a.in_place and os.path.abspath(a.out) == os.path.abspath(a.baseline):
        ap.error('--out must be a different file from --baseline (the baseline is never overwritten)')

    gs = None
    if a.gsheet_id:
        gs = open_gsheet(a.gsheet_id, a.gsheet_creds)
        wb, base_ids, surveyors = None, *read_gsheet(gs)
        if a.revalidate:
            print('(--revalidate is only available with --baseline; skipped)'); a.revalidate = False
    else:
        wb, base_ids, surveyors = read_baseline(a.baseline)
    print('Baseline: %d surveys, highest legacySurveyId %d' % (len(base_ids), max(base_ids)))

    # ---- load source
    if a.mysql:
        cn, cur, src_ids = load_mysql(a.prefix, None)
        new_ids = sorted(src_ids - base_ids)
        entries, metas = fetch_mysql_entries(cur, a.prefix, new_ids + (sorted(src_ids & base_ids)
                                                                      if (a.revalidate or a.all_media) else []))
    elif a.dump_dir:
        entries, metas = load_dirs(a.dump_dir)
        src_ids = set(entries)
        new_ids = sorted(src_ids - base_ids)
    else:
        if not (a.entry_csv and a.meta_csv):
            ap.error('give --mysql, --dump-dir, or both --entry-csv and --meta-csv')
        entries, metas = load_csv(a.entry_csv, a.meta_csv)
        src_ids = set(entries)
        new_ids = sorted(src_ids - base_ids)
    # ---- pre-flight
    bad_status = [i for i in new_ids if (entries[i].get('status') or 'active') != 'active']
    new_ids = [i for i in new_ids if i not in bad_status]
    missing_in_source = sorted(base_ids - src_ids)
    print('\n=== PRE-FLIGHT ===')
    print('Entries of form %d in source        : %d' % (FORM_ID, len(src_ids)))
    print('NEW entries (id not in baseline)    : %d%s' % (len(new_ids), ('  (IDs %d .. %d)' % (new_ids[0], new_ids[-1])
                                                                    if new_ids else '')))
    if bad_status:
        print('Skipped non-active entries          : %d  %s' % (len(bad_status), bad_status[:20]))
    print('Baseline IDs absent from the source : %d%s' % (len(missing_in_source),
                                                           ('  ' + str(missing_in_source[:20])) if missing_in_source else ''))
    built = {i: build(i, metas.get(i, []), entries[i], surveyors) for i in new_ids}
    hard = collections.defaultdict(list)
    HARD = {'BUSINESS_DEVELOPER_MISSING', 'INVALID_LEGACY_DATE', 'MISSING_COORDINATES', 'FLOOR_NAME_UNMAPPED',
            'NEW_DUPLICATE_META_KEYS', 'NEW_NEW_SURVEYOR', 'NEW_SURVEYOR_EMAIL', 'NEW_CITY_NO_ZIP',
            'PLOT_DIMENSIONS_UNPARSEABLE', 'NEW_MEDIA_NO_IMAGE', 'INVALID_NAME', 'COORDINATES_OUTSIDE_PAKISTAN'}
    for i, b in built.items():
        for e in b['exceptions']:
            if e[5] in HARD:
                hard[i].append(e[5])
    print('Entries with at least one field that cannot be mapped reliably: %d' % len(hard))
    byreason = collections.Counter(r for v in hard.values() for r in v)
    for r, c in byreason.most_common():
        print('   %-32s %d entries' % (r, c))
    if hard:
        print('   first affected IDs:', sorted(hard)[:25])
    plan_rows, plan_changes, plan_stats, plan_untouched = [], [], collections.Counter(), set()
    if ALL_MEDIA:
        plan_rows, plan_changes, plan_stats, plan_untouched = plan_all_media(wb, entries, metas, base_ids)
        new_media = sum(len(b['attachments']) for b in built.values())
        print('\n=== ALL MEDIA (fields 83 + 103 + 129) ===')
        print('Surveys already in the workbook : files unchanged %d, NEW %d, REMOVED (gone from the entry) %d, '
              'surveys kept as they are %d' % (plan_stats['unchanged'], plan_stats['added'], plan_stats['removed'],
                                               plan_stats['surveys_kept_as_is']))
        print('New surveys                     : %d media files' % new_media)
        print('04_Attachments after the run    : %d rows' % (len(plan_rows) + new_media))
        if not a.check_urls:
            print('(no --check-urls: new files get no size and dead links are NOT filtered out; run with --check-urls)')
    if a.dry_run:
        print('\n--dry-run: nothing written.')
        return

    # ---- revalidate rules on entries already in the baseline
    if a.revalidate:
        revalidate(wb, entries, metas, base_ids, surveyors)

    # ---- optional HEAD checks
    dropped_media = []                    # (survey id, row, reason): files that did not verify and were left out
    if a.check_urls and ALL_MEDIA:
        new_rows_all = [x for b in built.values() for x in b['attachments']]
        failed = check_media_rows(plan_rows + new_rows_all)
        bad_ids = {id(r) for r, _ in failed}
        plan_rows = [r for r in plan_rows if id(r) not in bad_ids]
        for b in built.values():
            b['attachments'] = [x for x in b['attachments'] if id(x) not in bad_ids]
        for r, reason in failed:
            sid = int(r['legacySurveyId'])
            dropped_media.append((sid, r, reason))
            if sid in built:
                built[sid]['exceptions'].append(('Attachment', sid, r['attachmentType'], r['originalUrl'], None,
                                                 'NEW_MEDIA_URL_NOT_REACHABLE', 'MEDIUM',
                                                 'Left out of 04_Attachments: %s' % reason))
    elif a.check_urls:
        head_check(built)

    # ---- write new workbook
    sheets = {'01_Surveys': [], '02_Survey_Floors': [], '03_Survey_Contacts': [], '04_Attachments': [],
              '05_Seed_Exceptions': []}
    for i in new_ids:
        b = built[i]
        sheets['01_Surveys'].append(b['survey'])
        sheets['02_Survey_Floors'] += b['floors']
        sheets['03_Survey_Contacts'].append(b['contact'])
        sheets['04_Attachments'] += b['attachments']
        sheets['05_Seed_Exceptions'] += [dict(zip(EXC_HDR, e)) for e in b['exceptions']]
    if gs is not None:
        gsheet_append(gs, sheets)
        write_summary(a, built, new_ids, hard, sheets)
        return
    if a.only_new_workbook:
        nw = openpyxl.Workbook(); nw.remove(nw.active)
        for name, rows in sheets.items():
            ws = nw.create_sheet(name)
            hdr = [c.value for c in wb[name][1]]
            ws.append(hdr)
            for r in rows:
                ws.append([r.get(h) for h in hdr])
                for c in ws[ws.max_row]:
                    if isinstance(c.value, datetime.datetime):
                        c.number_format = DATE_FMT
        nw.save(a.only_new_workbook)
    if ALL_MEDIA:
        source_active = {i for i in (src_ids | base_ids) if i in entries and i not in plan_untouched
                         and (entries[i].get('status') or 'active') == 'active' and i in metas}
        doc_rows, old_removed = refresh_media_exceptions(wb['05_Seed_Exceptions'], source_active & base_ids,
                                                         entries, metas)
        sheets['05_Seed_Exceptions'] += doc_rows
        sheets['05_Seed_Exceptions'] += [dict(zip(EXC_HDR, ('Attachment', sid, r['attachmentType'], r['originalUrl'],
                                                            None, 'NEW_MEDIA_URL_NOT_REACHABLE', 'MEDIUM',
                                                            'Left out of 04_Attachments: %s' % reason)))
                                         for sid, r, reason in dropped_media if sid not in built]
        merged = sorted(plan_rows + sheets['04_Attachments'], key=lambda r: int(r['legacySurveyId']))
        replace_rows(wb['04_Attachments'], merged)
        sheets['04_Attachments'] = merged
        print('04_Attachments rewritten: %d rows. 05_Seed_Exceptions: %d outdated media notes removed, %d backup-document '
              'notes written' % (len(merged), old_removed, len(doc_rows)))
        for name, rows in sheets.items():
            if name != '04_Attachments':
                append_rows(wb[name], rows)
    else:
        for name, rows in sheets.items():
            append_rows(wb[name], rows)
    wb.save(a.out)
    if a.in_place:
        import shutil
        bak = '%s.backup-%s.xlsx' % (os.path.splitext(a.baseline)[0], datetime.datetime.now().strftime('%Y%m%d-%H%M%S'))
        shutil.copy2(a.baseline, bak)
        os.replace(a.out, a.baseline)
        print('Backup of the previous file: %s' % bak)
        a.out = a.baseline
    # ---- summary
    if ALL_MEDIA:
        over = [r for r in sheets['04_Attachments']
                if (r.get('fileSizeBytes') or 0) > (IMAGE_LIMIT_BYTES if r['attachmentType'] == 'IMAGE'
                                                    else VIDEO_LIMIT_BYTES)]
        base_path = os.path.splitext(a.out)[0]
        changes_path = base_path + '_attachment_changes.tsv'
        with open(changes_path, 'w', encoding='utf-8', newline='') as f:
            f.write('change\tlegacySurveyId\ttype\toldUrl\tnewUrl\tnote\n')
            for c in plan_changes:
                f.write('\t'.join(str(x) for x in c) + '\n')
            for sid, r, reason in dropped_media:
                f.write('\t'.join(str(x) for x in ('NOT_ADDED', sid, r['attachmentType'], '', r['originalUrl'],
                                                  reason)) + '\n')
            for r in over:
                f.write('\t'.join(str(x) for x in ('OVER_APP_SIZE_LIMIT', r['legacySurveyId'], r['attachmentType'], '',
                                                  r['originalUrl'], '%d bytes (app limit %s)' % (
                                                      r['fileSizeBytes'],
                                                      '10 MB' if r['attachmentType'] == 'IMAGE' else '300 MB'))) + '\n')
        ALL_MEDIA_REPORT.update(
            files_in_04_attachments=len(sheets['04_Attachments']),
            surveys_with_media=len({int(r['legacySurveyId']) for r in sheets['04_Attachments']}),
            images=sum(1 for r in sheets['04_Attachments'] if r['attachmentType'] == 'IMAGE'),
            videos=sum(1 for r in sheets['04_Attachments'] if r['attachmentType'] == 'VIDEO'),
            existing_files_unchanged=plan_stats['unchanged'], existing_surveys_new_files=plan_stats['added'],
            existing_files_removed=plan_stats['removed'], surveys_kept_as_they_are=plan_stats['surveys_kept_as_is'],
            new_files_not_added_unreachable=len(dropped_media), files_over_app_size_limit=len(over),
            urls_checked=bool(a.check_urls), changes_file=os.path.basename(changes_path))
    write_summary(a, built, new_ids, hard, sheets)


def head_check(built):
    import urllib.request
    n = 0
    for i, b in built.items():
        for x in b['attachments']:
            try:
                req = urllib.request.Request(x['originalUrl'], method='HEAD')
                with urllib.request.urlopen(req, timeout=20) as r:
                    cl = r.headers.get('Content-Length')
                    x['fileSizeBytes'] = int(cl) if cl and cl.isdigit() else None
            except Exception as ex:
                code = getattr(ex, 'code', None)
                b['exceptions'].append(('Attachment', i, x['attachmentType'], x['originalUrl'], None,
                                        'NEW_MEDIA_URL_NOT_REACHABLE', 'MEDIUM',
                                        'HEAD request failed (%s)' % (code or type(ex).__name__)))
            n += 1
    print('HEAD-checked %d media URLs (headers only, nothing downloaded)' % n)


def write_summary(a, built, new_ids, hard, sheets):
    exc = collections.Counter(e['reason'] for e in sheets['05_Seed_Exceptions'])
    sev = collections.Counter(e['severity'] for e in sheets['05_Seed_Exceptions'])
    media = collections.Counter()
    for r in sheets['05_Seed_Exceptions']:
        if r['reason'].startswith('NEW_MEDIA'):
            media[r['reason']] += 1
    ids_no_img = sorted({r['legacySurveyId'] for r in sheets['05_Seed_Exceptions'] if r['reason'] == 'NEW_MEDIA_NO_IMAGE'})
    summ = collections.OrderedDict()
    summ['new_surveys'] = len(new_ids)
    summ['id_range'] = [new_ids[0], new_ids[-1]] if new_ids else None
    summ['floor_rows'] = len(sheets['02_Survey_Floors'])
    summ['contact_rows'] = len(sheets['03_Survey_Contacts'])
    summ['attachment_rows'] = len(sheets['04_Attachments'])
    summ['exception_rows'] = len(sheets['05_Seed_Exceptions'])
    summ['exceptions_by_reason'] = dict(exc.most_common())
    summ['exceptions_by_severity'] = dict(sev)
    summ['entries_not_reliably_mappable'] = len(hard)
    summ['media_needing_attention'] = dict(media)
    summ['surveys_without_any_image'] = len(ids_no_img)
    if ALL_MEDIA:
        summ['attachment_rows'] = len(sheets['04_Attachments'])        # whole tab (all surveys), not only new rows
        summ['all_media'] = dict(ALL_MEDIA_REPORT)
    path =a.summary or ((os.path.splitext(a.out)[0] if a.out else 'gsheet_sync') + '_summary.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(summ, f, indent=2)
    print('\n=== RESULT ===')
    print('Wrote workbook : %s' % (a.out or 'Google Sheet ' + str(a.gsheet_id)))
    print('Wrote summary  : %s' % path)
    print(json.dumps(summ, indent=2))
    if ALL_MEDIA:
        print('\n--all-media: existing survey, floor and contact rows were not modified. 04_Attachments was rebuilt '
              '(unchanged files kept as they are) and the outdated media notes in 05_Seed_Exceptions were refreshed; '
              'see %s.' % ALL_MEDIA_REPORT.get('changes_file'))
        if not ALL_MEDIA_REPORT.get('urls_checked'):
            print('NOTE: --check-urls was not used, so new files have no size and dead links were not filtered out.')
        print('Nothing was written to any database.')
    else:
        print('\nExisting rows were not modified (append only). Nothing was written to any database.')


def revalidate(wb, entries, metas, base_ids, surveyors):
    hdr = [c.value for c in wb['01_Surveys'][1]]
    rows = {int(r[0]): dict(zip(hdr, r)) for r in wb['01_Surveys'].iter_rows(min_row=2, values_only=True) if r[0]}
    mism = collections.Counter(); n = 0
    cols = ['address', 'latitude', 'longitude', 'city', 'area', 'zipCode', 'propertyType', 'ownerName', 'landArea',
            'landAreaUnit', 'grossBuildingArea', 'numberOfFloors', 'propertyStatus', 'availableFor', 'advanceAmount',
            'securityDepositAmount', 'possessionTerms', 'surveyDate', 'neighboringBusinesses', 'createdById',
            'createdAt']
    for i in sorted(set(entries) & base_ids):
        if i not in metas:
            continue
        n += 1
        s = build(i, metas[i], entries[i], surveyors)['survey']
        for c in cols:
            if c not in hdr:
                continue
            if s.get(c) != rows[i].get(c):
                mism[c] += 1
    print('\n=== REVALIDATE: %d baseline entries re-derived from the source ===' % n)
    for c in cols:
        if c in hdr:
            print('  %-24s differs in %d of %d' % (c, mism[c], n))
    print('  (differences = entry edited at the source since the baseline was built, or a limit of the rule;\n'
          '   neighboringBusinesses is free text and the least exact column)')


if __name__ == '__main__':
    main()
