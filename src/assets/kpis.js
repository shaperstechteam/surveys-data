/**
 * Compute all dashboard KPIs from a list of property records.
 * Records use compact keys from properties.json (ci, on, af, am, etc.)
 */

function parseAmount(v) {
  if (v == null || v === '') return null
  const n = Number(String(v).replace(/,/g, ''))
  return Number.isFinite(n) ? n : null
}

function isFilled(v) {
  return v != null && String(v).trim() !== ''
}

function countBy(rows, key, { limit = 0, normalize = (x) => x } = {}) {
  const map = new Map()
  for (const r of rows) {
    let v = normalize(r[key])
    if (v == null || v === '' || v === '.') v = '(empty)'
    map.set(v, (map.get(v) || 0) + 1)
  }
  let list = [...map.entries()]
    .map(([name, count]) => ({
      name,
      count,
      pct: rows.length ? (count / rows.length) * 100 : 0,
    }))
    .sort((a, b) => b.count - a.count)
  if (limit > 0) list = list.slice(0, limit)
  return list
}

function countMulti(rows, key, { limit = 20 } = {}) {
  // Select_Brands may be JSON array string or comma-separated
  const map = new Map()
  let filled = 0
  for (const r of rows) {
    const raw = r[key]
    if (!isFilled(raw)) continue
    filled++
    let items = []
    const s = String(raw).trim()
    if (s.startsWith('[')) {
      try {
        const parsed = JSON.parse(s)
        if (Array.isArray(parsed)) items = parsed.map(String)
      } catch {
        items = s.split(/[,;|]/).map((x) => x.trim()).filter(Boolean)
      }
    } else {
      items = s.split(/[,;|]/).map((x) => x.trim()).filter(Boolean)
    }
    for (const item of items) {
      const name = item.replace(/^["']|["']$/g, '').trim()
      if (!name) continue
      map.set(name, (map.get(name) || 0) + 1)
    }
  }
  const list = [...map.entries()]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, limit)
  return { list, filled, total: rows.length }
}

function median(nums) {
  if (!nums.length) return null
  const a = [...nums].sort((x, y) => x - y)
  const m = Math.floor(a.length / 2)
  return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2
}

function amountBands(amounts) {
  const bands = [
    { name: '< 1 Lakh', min: 0, max: 100000 },
    { name: '1 – 5 Lakh', min: 100000, max: 500000 },
    { name: '5 – 10 Lakh', min: 500000, max: 1000000 },
    { name: '10 – 25 Lakh', min: 1000000, max: 2500000 },
    { name: '25 – 50 Lakh', min: 2500000, max: 5000000 },
    { name: '50 Lakh+', min: 5000000, max: Infinity },
  ]
  return bands.map((b) => {
    const count = amounts.filter((n) => n >= b.min && n < b.max).length
    return {
      name: b.name,
      count,
      pct: amounts.length ? (count / amounts.length) * 100 : 0,
    }
  })
}

function monthKey(dc) {
  if (!dc) return null
  const s = String(dc).trim()
  // expect YYYY-MM-DD or similar
  const m = s.match(/^(\d{4})-(\d{2})/)
  if (!m) return null
  if (m[1] === '0000') return null
  return `${m[1]}-${m[2]}`
}

/**
 * @param {Array<Record<string, any>>} rows - filtered records
 * @returns {object} kpi payload
 */
export function computeKpis(rows) {
  const total = rows.length

  const amounts = rows.map((r) => parseAmount(r.am)).filter((n) => n != null)
  const withAmount = amounts.length
  const withPic = rows.filter((r) => isFilled(r.pic)).length
  const withVid = rows.filter((r) => isFilled(r.vid)).length
  const withCoords = rows.filter((r) => isFilled(r.lng) && isFilled(r.lat)).length
  const missingPhone = rows.filter((r) => !isFilled(r.cn)).length
  const missingOwner = rows.filter((r) => !isFilled(r.on)).length
  const missingCity = rows.filter((r) => !isFilled(r.ci) || r.ci === '.').length
  const missingAmount = total - withAmount
  const withAdvance = rows.filter((r) => isFilled(r.av)).length
  const withSecurity = rows.filter((r) => isFilled(r.sc)).length
  const withBrands = rows.filter(
    (r) => isFilled(r.sb) || isFilled(r.eb) || isFilled(r.ob),
  ).length

  const owners = new Map()
  for (const r of rows) {
    const name = isFilled(r.on) ? String(r.on).trim() : null
    if (!name) continue
    owners.set(name, (owners.get(name) || 0) + 1)
  }
  const uniqueOwners = owners.size
  const topOwners = [...owners.entries()]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 15)

  const byCity = countBy(rows, 'ci', { limit: 25 })
  const byAvailFor = countBy(rows, 'af')
  const byNotRented = countBy(rows, 'an')
  const byPropType = countBy(rows, 'pt')
  const byDeal = countBy(rows, 'dl')
  const byBd = countBy(rows, 'bd', {
    limit: 30,
    normalize: (v) => {
      if (!isFilled(v)) return '(empty)'
      // show email local-part or full
      return String(v).trim()
    },
  })

  // Areas (city-specific)
  const areasLahore = countBy(
    rows.filter((r) => r.ci === 'Lahore'),
    'al',
    { limit: 15 },
  )
  const areasIslamabad = countBy(
    rows.filter((r) => r.ci === 'Islamabad'),
    'ai',
    { limit: 15 },
  )
  const areasRawalpindi = countBy(
    rows.filter((r) => r.ci === 'Rawalpindi'),
    'ar',
    { limit: 15 },
  )
  const areasKarachi = countBy(
    rows.filter((r) => r.ci === 'Karachi'),
    'ak',
    { limit: 15 },
  )

  const brands = countMulti(rows, 'sb', { limit: 25 })

  // BD × City (top pairs)
  const bdCity = new Map()
  for (const r of rows) {
    if (!isFilled(r.bd) || !isFilled(r.ci)) continue
    const k = `${r.bd}|||${r.ci}`
    bdCity.set(k, (bdCity.get(k) || 0) + 1)
  }
  const bdCityTop = [...bdCity.entries()]
    .map(([k, count]) => {
      const [bd, city] = k.split('|||')
      return { bd, city, count }
    })
    .sort((a, b) => b.count - a.count)
    .slice(0, 20)

  // Completeness %
  const fields = [
    { key: 'on', label: 'Owner Name' },
    { key: 'cn', label: 'Contact Number' },
    { key: 'ci', label: 'City' },
    { key: 'ad', label: 'Address' },
    { key: 'am', label: 'Amount' },
    { key: 'ta', label: 'Total Area' },
    { key: 'af', label: 'Available For' },
    { key: 'pic', label: 'Pictures' },
    { key: 'vid', label: 'Videos' },
    { key: 'bd', label: 'Business Developer' },
    { key: 'sb', label: 'Select Brands' },
    { key: 'lng', label: 'Coordinates' },
  ]
  const completeness = fields.map(({ key, label }) => {
    let count
    if (key === 'lng') count = withCoords
    else count = rows.filter((r) => isFilled(r[key])).length
    return {
      label,
      count,
      pct: total ? (count / total) * 100 : 0,
    }
  })

  // Entries by month
  const byMonthMap = new Map()
  for (const r of rows) {
    const k = monthKey(r.dc)
    if (!k) continue
    byMonthMap.set(k, (byMonthMap.get(k) || 0) + 1)
  }
  const byMonth = [...byMonthMap.entries()]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => a.name.localeCompare(b.name))

  const avg = amounts.length
    ? amounts.reduce((s, n) => s + n, 0) / amounts.length
    : null
  const med = median(amounts)
  const minA = amounts.length ? Math.min(...amounts) : null
  const maxA = amounts.length ? Math.max(...amounts) : null

  // Amount by city (avg for top cities)
  const cityAmounts = new Map()
  for (const r of rows) {
    const n = parseAmount(r.am)
    if (n == null || !isFilled(r.ci)) continue
    if (!cityAmounts.has(r.ci)) cityAmounts.set(r.ci, [])
    cityAmounts.get(r.ci).push(n)
  }
  const amountByCity = [...cityAmounts.entries()]
    .map(([city, arr]) => ({
      city,
      count: arr.length,
      avg: arr.reduce((s, n) => s + n, 0) / arr.length,
      median: median(arr),
      min: Math.min(...arr),
      max: Math.max(...arr),
    }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 15)

  return {
    total,
    withAmount,
    withPic,
    withVid,
    withCoords,
    missingPhone,
    missingOwner,
    missingCity,
    missingAmount,
    withAdvance,
    withSecurity,
    withBrands,
    uniqueOwners,
    topOwners,
    byCity,
    byAvailFor,
    byNotRented,
    byPropType,
    byDeal,
    byBd,
    areasLahore,
    areasIslamabad,
    areasRawalpindi,
    areasKarachi,
    brands,
    bdCityTop,
    completeness,
    byMonth,
    amount: {
      count: withAmount,
      avg,
      median: med,
      min: minA,
      max: maxA,
      bands: amountBands(amounts),
      byCity: amountByCity,
    },
  }
}

export function formatNumber(n) {
  if (n == null || Number.isNaN(n)) return '—'
  return Number(n).toLocaleString()
}

export function formatMoney(n) {
  if (n == null || Number.isNaN(n)) return '—'
  return Number(n).toLocaleString(undefined, { maximumFractionDigits: 0 })
}

export function formatPct(n) {
  if (n == null || Number.isNaN(n)) return '—'
  return `${n.toFixed(1)}%`
}
