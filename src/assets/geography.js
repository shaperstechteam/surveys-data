// Geographic split uses the survey City field (`ci`). That field is filled
// on almost every record — unlike sector, this is real form data, not a guess.

export const OTHER_CITIES = '__other__'

export const GEO_CITIES = [
  {
    key: 'Lahore',
    label: 'Lahore',
    description: 'Punjab',
    color: '#3b82f6',
  },
  {
    key: 'Islamabad',
    label: 'Islamabad',
    description: 'Federal capital',
    color: '#06b6d4',
  },
  {
    key: 'Rawalpindi',
    label: 'Rawalpindi',
    description: 'Punjab',
    color: '#8b5cf6',
  },
  {
    key: 'Karachi',
    label: 'Karachi',
    description: 'Sindh',
    color: '#f59e0b',
  },
  {
    key: 'Faisalabad',
    label: 'Faisalabad',
    description: 'Punjab',
    color: '#22c55e',
  },
  {
    key: 'Peshawar',
    label: 'Peshawar',
    description: 'Khyber Pakhtunkhwa',
    color: '#ef4444',
  },
  {
    key: OTHER_CITIES,
    label: 'Other cities',
    description: 'All remaining cities',
    color: '#64748b',
  },
]

export const MAJOR_CITY_KEYS = GEO_CITIES
  .map((g) => g.key)
  .filter((k) => k !== OTHER_CITIES)

export function matchesCityFilter(recordCity, selected) {
  if (!selected) return true
  const ci = recordCity || ''
  if (selected === OTHER_CITIES) return !MAJOR_CITY_KEYS.includes(ci)
  return ci === selected
}
