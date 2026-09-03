// Geographic/use-type segregation of Shapers surveys.
//
// The raw data has no explicit "sector" field, so each record is classified
// by scanning its text fields (nearby/earlier brands, address, notes,
// property type) for keywords associated with each sector. This is a
// heuristic best-effort classification, not ground truth — records with no
// matching signal fall back to "Retail" (Shapers' core business: shops on
// malls/high streets) or "Uncategorized" when there's no usable text at all.

export const SECTORS = [
  {
    key: 'retail',
    label: 'Retail',
    description: 'Malls, high streets, supermarkets',
    color: '#3b82f6',
  },
  {
    key: 'fnb',
    label: 'F&B',
    description: 'Restaurants, cafés, food courts',
    color: '#f97316',
  },
  {
    key: 'corporate',
    label: 'Corporate',
    description: 'Offices, headquarters, business parks',
    color: '#8b5cf6',
  },
  {
    key: 'healthcare',
    label: 'Healthcare',
    description: 'Hospitals, clinics, pharmacies',
    color: '#ef4444',
  },
  {
    key: 'education',
    label: 'Education',
    description: 'Schools, universities, training centres',
    color: '#eab308',
  },
  {
    key: 'hospitality',
    label: 'Hospitality',
    description: 'Hotels, serviced apartments',
    color: '#14b8a6',
  },
  {
    key: 'industrial',
    label: 'Industrial',
    description: 'Warehouses, logistics, manufacturing',
    color: '#78716c',
  },
  {
    key: 'uncategorized',
    label: 'Uncategorized',
    description: 'Not enough data to classify',
    color: '#64748b',
  },
]

export const SECTOR_MAP = Object.fromEntries(SECTORS.map((s) => [s.key, s]))

// Keyword lists, checked in priority order (most specific/rare first) so a
// hospital pharmacy or a hotel restaurant lands in the more distinctive
// sector rather than the generic Retail bucket. Each entry may be a plain
// substring or, where word-boundary precision matters, a RegExp.
const RULES = [
  {
    key: 'healthcare',
    keywords: [
      'hospital', 'clinic', 'pharmacy', 'pharmacia', 'chemist', 'dispensary',
      'dispensing', 'medical', 'diagnostic', 'laboratory', /\blab\b/, 'dental',
      'dentist', 'healthcare', 'health care', 'maternity', 'poly clinic',
      'polyclinic', 'blood bank', 'surgical', 'physiotherapy', 'dawakhana',
    ],
  },
  {
    key: 'education',
    keywords: [
      'school', 'college', 'university', 'academy', 'institute', 'campus',
      'training center', 'training centre', 'montessori', 'madrassa',
      'madrasa', 'grammar school', 'public school', 'tuition center',
      'tuition centre',
    ],
  },
  {
    key: 'hospitality',
    keywords: [
      'hotel', 'motel', 'resort', 'guest house', 'guesthouse', 'inn ',
      'serviced apartment', 'rest house', 'resthouse',
    ],
  },
  {
    key: 'industrial',
    keywords: [
      'warehouse', 'godown', 'go down', 'factory', 'industrial estate',
      'industrial area', 'industrial zone', 'logistics', 'cold storage',
      'manufacturing', /\bmill\b/, /\bplant\b/, 'production unit',
    ],
  },
  {
    key: 'fnb',
    keywords: [
      // Note: 'diner'/"diner's" deliberately excluded — it collides with
      // "Diners", a menswear retail brand that dominates the brand lists.
      'restaurant', 'resturant', 'café', 'cafe', 'food court', 'foodcourt',
      'bakery', 'bakers', 'sweets', 'tikka', 'bar b', 'bbq', 'broast',
      'biryani', 'dhaba', 'kabab', 'fast food', 'ice cream', 'kfc',
      "mcdonald", 'dominos', "domino's", 'pizza hut', 'broadway pizza',
      'bundu khan', 'subway', 'hardee', 'nando', 'cheezious', 'gloria jeans',
      'second cup', 'fri chicks', 'bombay chowpatti', 'gourmet bakers',
      'cakes & bakes', 'coffee planet', 'yum chinese', 'burger king',
      'karachi broast', 'china town', 'zubaida', 'coffee', 'pizza',
      'chicken', 'optp', 'baskin robbins', 'dunkin donuts', 'cbtl',
      'asian wok', 'burger lab',
    ],
  },
  {
    key: 'corporate',
    keywords: [
      'head office', 'headquarters', 'headquarter', 'business park',
      'corporate office', 'corporate tower', 'software house', 'call center',
      'call centre', 'it park', 'i.t. park',
    ],
  },
]

// 'sb' ("Select brands") is intentionally excluded: it's a mostly-templated
// checklist of generic retail brand names reused near-verbatim across
// thousands of records, so it signals "this is a plain retail unit" rather
// than anything distinctive — including it would just drown out real
// signal from the address/notes fields with retail-brand noise.
function extractText(record) {
  const fields = [
    'ad', 'ak', 'ai', 'ar', 'al', 'm4', 'm7', 'ms', 'eb', 'ob', 'no',
    'nt', 'ta',
  ]
  return fields
    .map((f) => record[f] || '')
    .join(' ')
    .toLowerCase()
}

function matches(text, keyword) {
  if (keyword instanceof RegExp) return keyword.test(text)
  return text.includes(keyword)
}

/**
 * Classify a property/survey record into one of the SECTORS keys.
 * Pure function of the record's text fields — safe to memoize per record.
 */
export function classifySector(record) {
  const text = extractText(record)
  const pt = (record.pt || '').toLowerCase()

  for (const rule of RULES) {
    if (rule.keywords.some((kw) => matches(text, kw))) return rule.key
  }

  // No distinctive keyword hit. An explicit "office" property type is a
  // strong direct signal for Corporate.
  if (pt === 'office') return 'corporate'

  // Nothing to go on at all — don't force a guess.
  if (!text.trim() && !pt) return 'uncategorized'

  // Default: Shapers' core business is retail leasing (shops on malls and
  // high streets), and the vast majority of unmatched records are fashion
  // and general-merchandise brands, so they fall back here.
  return 'retail'
}

export function sectorLabel(key) {
  return SECTOR_MAP[key]?.label || key
}
