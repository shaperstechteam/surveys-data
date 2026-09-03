import { useCallback, useEffect, useMemo, useState } from 'react'
import './App.css'
import { computeKpis } from './assets/kpis'
import { SECTORS, SECTOR_MAP, classifySector } from './assets/sectors'
import KpiDashboard from './components/KpiDashboard'

const PAGE_SIZE = 50

const LABELS = {
  id: 'ID',
  dc: 'Date created',
  st: 'Status',
  on: 'Owner name',
  cp: 'Contact person',
  cn: 'Contact number',
  bd: 'Business developer',
  ad: 'Address',
  ci: 'City',
  al: 'Area (Lahore)',
  ai: 'Area (Islamabad)',
  ar: 'Area (Rawalpindi)',
  ak: 'Area (Karachi)',
  m4: 'Market area',
  m7: 'Market area 2',
  lng: 'Longitude',
  lat: 'Latitude',
  ms: 'Map search address',
  pt: 'Property type',
  ps: 'Plot size',
  pa: 'Plot area',
  cs: 'Covered size',
  sf: 'Select floors',
  ta: 'Total area',
  af: 'Available for',
  rp: 'Rent/Sale price available',
  nt: 'Price note',
  am: 'Amount',
  dl: 'Deal',
  av: 'Advance',
  sc: 'Security',
  po: 'Possession',
  gp: 'Grace period',
  inc: 'Increment',
  ap: 'Agreement period',
  no: 'Note',
  eb: 'Earlier brands',
  sb: 'Select brands',
  ob: 'Other brands',
  an: 'Available not rented',
  nf: 'No. of floors',
  f1t: 'Floor 1 type',
  f1a: 'Floor 1 area',
  f1s: 'Floor 1 size',
  f2t: 'Floor 2 type',
  f2a: 'Floor 2 area',
  f2s: 'Floor 2 size',
  f3t: 'Floor 3 type',
  f3a: 'Floor 3 area',
  f3s: 'Floor 3 size',
  pic: 'Picture',
  vid: 'Video',
}

const SECTIONS = [
  { title: 'Contact', keys: ['on', 'cp', 'cn', 'bd'] },
  {
    title: 'Location',
    keys: ['ad', 'ci', 'al', 'ai', 'ar', 'ak', 'm4', 'm7', 'ms', 'lng', 'lat'],
  },
  {
    title: 'Property',
    keys: [
      'pt', 'ps', 'pa', 'cs', 'sf', 'ta', 'nf',
      'f1t', 'f1a', 'f1s', 'f2t', 'f2a', 'f2s', 'f3t', 'f3a', 'f3s',
    ],
  },
  {
    title: 'Deal / Pricing',
    keys: ['af', 'rp', 'nt', 'am', 'dl', 'av', 'sc', 'po', 'gp', 'inc', 'ap', 'an'],
  },
  { title: 'Brands & Notes', keys: ['eb', 'sb', 'ob', 'no'] },
  { title: 'Meta', keys: ['dc', 'st', 'id'] },
]

function formatAmount(v) {
  if (!v) return ''
  const n = Number(String(v).replace(/,/g, ''))
  if (Number.isNaN(n)) return String(v)
  return n.toLocaleString()
}

function TypeBadge({ value }) {
  const v = (value || '').toLowerCase()
  if (v.includes('sale')) return <span className="badge sale">Sale</span>
  if (v.includes('rent')) return <span className="badge rent">Rent</span>
  return <span className="badge empty">{value || '—'}</span>
}

function SectorBadge({ sectorKey }) {
  const sector = SECTOR_MAP[sectorKey]
  if (!sector) return <span className="badge empty">—</span>
  return (
    <span
      className="badge"
      style={{ background: `${sector.color}26`, color: sector.color }}
      title={sector.description}
    >
      {sector.label}
    </span>
  )
}

function DetailPanel({ record, onClose }) {
  if (!record) {
    return (
      <aside className="detail-panel">
        <div className="detail-header">
          <h2>Select a record</h2>
        </div>
        <div className="detail-body">
          <div className="empty-state">
            Click any row to view full details with all fields.
          </div>
        </div>
      </aside>
    )
  }

  return (
    <aside className="detail-panel open">
      <div className="detail-header">
        <div>
          <h2>
            #{record.id} — {record.on || 'Unknown owner'}
          </h2>
          <div className="detail-sub">
            <SectorBadge sectorKey={record._sector} />{' '}
            {[record.ci, record.af, record.st].filter(Boolean).join(' · ')}
          </div>
        </div>
        <button type="button" className="close-btn" onClick={onClose}>
          ✕ Close
        </button>
      </div>
      <div className="detail-body">
        {record.pic ? (
          <img
            className="pic-preview"
            src={record.pic}
            alt="Property"
            onError={(e) => {
              e.currentTarget.style.display = 'none'
            }}
          />
        ) : null}
        {record.vid ? (
          <div className="field" style={{ marginBottom: 12 }}>
            <label>Video</label>
            <div className="val">
              <a href={record.vid} target="_blank" rel="noopener noreferrer">
                Open video
              </a>
            </div>
          </div>
        ) : null}

        {SECTIONS.map((sec) => {
          const fields = sec.keys.filter((k) => record[k])
          if (!fields.length) return null
          return (
            <div key={sec.title}>
              <div className="section-title">{sec.title}</div>
              <div className="grid2">
                {fields.map((k) => {
                  let val = record[k]
                  if (k === 'am') val = formatAmount(val) || val
                  return (
                    <div className="field" key={k}>
                      <label>{LABELS[k] || k}</label>
                      <div className="val">{val}</div>
                    </div>
                  )
                })}
              </div>
            </div>
          )
        })}

        {record.lng && record.lat ? (
          <>
            <div className="section-title">Map</div>
            <div className="field">
              <label>Coordinates</label>
              <div className="val">
                {record.lat}, {record.lng} ·{' '}
                <a
                  href={`https://www.google.com/maps?q=${record.lat},${record.lng}`}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Open in Google Maps
                </a>
              </div>
            </div>
          </>
        ) : null}
      </div>
    </aside>
  )
}

export default function App() {
  const [all, setAll] = useState([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const [query, setQuery] = useState('')
  const [city, setCity] = useState('')
  const [sector, setSector] = useState('')
  const [avail, setAvail] = useState('')
  const [hasPrice, setHasPrice] = useState('')
  const [page, setPage] = useState(0)
  const [sortKey, setSortKey] = useState('id')
  const [sortDir, setSortDir] = useState(-1)
  const [selectedId, setSelectedId] = useState(null)
  const [showKpis, setShowKpis] = useState(true)

  useEffect(() => {
    let cancelled = false
    ;(async () => {
      try {
        const res = await fetch('/properties.json')
        if (!res.ok) throw new Error(`HTTP ${res.status}`)
        const data = await res.json()
        if (!cancelled) {
          setAll(data.map((r) => ({ ...r, _sector: classifySector(r) })))
          setLoading(false)
        }
      } catch (e) {
        if (!cancelled) {
          setError(e.message || 'Failed to load data')
          setLoading(false)
        }
      }
    })()
    return () => {
      cancelled = true
    }
  }, [])

  const cities = useMemo(() => {
    const set = new Set()
    all.forEach((r) => {
      if (r.ci && r.ci !== '.') set.add(r.ci)
    })
    return [...set].sort()
  }, [all])

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase()
    let rows = all.filter((r) => {
      if (city && r.ci !== city) return false
      if (sector && r._sector !== sector) return false
      if (avail && !(r.af || '').includes(avail)) return false
      if (hasPrice === 'yes' && !r.am) return false
      if (hasPrice === 'no' && r.am) return false
      if (q && !(r._s || '').includes(q)) return false
      return true
    })

    rows = [...rows].sort((a, b) => {
      let va = a[sortKey] ?? ''
      let vb = b[sortKey] ?? ''
      if (sortKey === 'id' || sortKey === 'am') {
        va = parseFloat(String(va).replace(/,/g, '')) || 0
        vb = parseFloat(String(vb).replace(/,/g, '')) || 0
      } else {
        va = String(va).toLowerCase()
        vb = String(vb).toLowerCase()
      }
      if (va < vb) return -1 * sortDir
      if (va > vb) return 1 * sortDir
      return 0
    })
    return rows
  }, [all, query, city, sector, avail, hasPrice, sortKey, sortDir])

  const kpis = useMemo(() => computeKpis(filtered), [filtered])
  const isFiltered = filtered.length !== all.length

  const sectorCounts = useMemo(() => {
    const counts = {}
    all.forEach((r) => {
      counts[r._sector] = (counts[r._sector] || 0) + 1
    })
    return counts
  }, [all])

  useEffect(() => {
    setPage(0)
  }, [query, city, sector, avail, hasPrice, sortKey, sortDir])

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE))
  const pageRows = filtered.slice(page * PAGE_SIZE, page * PAGE_SIZE + PAGE_SIZE)
  const selected = all.find((r) => String(r.id) === String(selectedId)) || null

  const onSort = useCallback(
    (key) => {
      if (sortKey === key) setSortDir((d) => d * -1)
      else {
        setSortKey(key)
        setSortDir(1)
      }
    },
    [sortKey],
  )

  const clearFilters = () => {
    setQuery('')
    setCity('')
    setSector('')
    setAvail('')
    setHasPrice('')
  }

  const exportCsv = () => {
    if (!filtered.length) {
      alert('No rows to export')
      return
    }
    const keys = Object.keys(LABELS)
    const header = ['Sector', ...keys.map((k) => LABELS[k] || k)].join(',')
    const rows = filtered.map((r) =>
      [SECTOR_MAP[r._sector]?.label || r._sector, ...keys.map((k) => r[k] ?? '')]
        .map((v) => {
          v = String(v).replace(/"/g, '""')
          if (/[",\n]/.test(v)) v = `"${v}"`
          return v
        })
        .join(','),
    )
    const blob = new Blob([[header, ...rows].join('\n')], {
      type: 'text/csv;charset=utf-8',
    })
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = `properties_filtered_${filtered.length}.csv`
    a.click()
  }

  if (loading) {
    return <div className="loading">Loading property records…</div>
  }

  if (error) {
    return (
      <div className="error">
        Failed to load data: {error}
        <br />
        Make sure <code>public/properties.json</code> exists and run{' '}
        <code>npm run dev</code>.
      </div>
    )
  }

  return (
    <div className="app">
      <header className="header">
        <h1>
          Property <span>Records</span> Viewer
        </h1>
        <div className="stats">
          Showing <strong>{filtered.length.toLocaleString()}</strong> of{' '}
          <strong>{all.length.toLocaleString()}</strong> records
        </div>
      </header>

      <div className="toolbar">
        <div className="search-wrap">
          <input
            type="search"
            placeholder="Search any field: owner, city, address, amount, brands, phone…"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            autoComplete="off"
          />
        </div>
        <select value={city} onChange={(e) => setCity(e.target.value)}>
          <option value="">All cities</option>
          {cities.map((c) => (
            <option key={c} value={c}>
              {c}
            </option>
          ))}
        </select>
        <select value={sector} onChange={(e) => setSector(e.target.value)}>
          <option value="">All sectors</option>
          {SECTORS.map((s) => (
            <option key={s.key} value={s.key}>
              {s.label} ({(sectorCounts[s.key] || 0).toLocaleString()})
            </option>
          ))}
        </select>
        <select value={avail} onChange={(e) => setAvail(e.target.value)}>
          <option value="">All types</option>
          <option value="Rent">Rent</option>
          <option value="Sale">Sale</option>
        </select>
        <select value={hasPrice} onChange={(e) => setHasPrice(e.target.value)}>
          <option value="">Price: any</option>
          <option value="yes">Has amount</option>
          <option value="no">No amount</option>
        </select>
        <button type="button" className="btn secondary" onClick={clearFilters}>
          Clear
        </button>
        <button type="button" className="btn secondary" onClick={exportCsv}>
          Export filtered CSV
        </button>
        <button
          type="button"
          className="btn secondary"
          onClick={() => setShowKpis((v) => !v)}
        >
          {showKpis ? 'Hide KPIs' : 'Show KPIs'}
        </button>
      </div>

      {showKpis ? <KpiDashboard kpis={kpis} filtered={isFiltered} /> : null}

      <main className="main">
        <div className="list-panel">
          <div className="pager">
            <button
              type="button"
              className="btn secondary sm"
              disabled={page <= 0}
              onClick={() => setPage((p) => Math.max(0, p - 1))}
            >
              ← Prev
            </button>
            <span>
              Page {page + 1} / {totalPages} ({filtered.length} rows)
            </span>
            <button
              type="button"
              className="btn secondary sm"
              disabled={page >= totalPages - 1}
              onClick={() => setPage((p) => p + 1)}
            >
              Next →
            </button>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th
                    className={sortKey === 'id' ? 'sorted' : ''}
                    onClick={() => onSort('id')}
                  >
                    ID ↕
                  </th>
                  <th
                    className={sortKey === 'ci' ? 'sorted' : ''}
                    onClick={() => onSort('ci')}
                  >
                    City ↕
                  </th>
                  <th
                    className={sortKey === 'on' ? 'sorted' : ''}
                    onClick={() => onSort('on')}
                  >
                    Owner ↕
                  </th>
                  <th>Address / Area</th>
                  <th>Sector</th>
                  <th>Type</th>
                  <th
                    className={sortKey === 'am' ? 'sorted' : ''}
                    onClick={() => onSort('am')}
                  >
                    Amount ↕
                  </th>
                  <th>Area</th>
                  <th>BD</th>
                </tr>
              </thead>
              <tbody>
                {pageRows.length === 0 ? (
                  <tr>
                    <td colSpan={9}>
                      <div className="empty-state">No records match your filters.</div>
                    </td>
                  </tr>
                ) : (
                  pageRows.map((r) => {
                    const area =
                      r.ad || r.al || r.ai || r.ar || r.ak || r.m4 || r.ms || '—'
                    const bd = (r.bd || '').split('@')[0] || '—'
                    return (
                      <tr
                        key={r.id}
                        className={
                          String(selectedId) === String(r.id) ? 'active' : ''
                        }
                        onClick={() => setSelectedId(r.id)}
                      >
                        <td>{r.id}</td>
                        <td title={r.ci}>{r.ci || '—'}</td>
                        <td title={r.on}>{r.on || '—'}</td>
                        <td title={area}>{area}</td>
                        <td>
                          <SectorBadge sectorKey={r._sector} />
                        </td>
                        <td>
                          <TypeBadge value={r.af} />
                        </td>
                        <td>{formatAmount(r.am) || '—'}</td>
                        <td title={r.ta}>{r.ta || '—'}</td>
                        <td title={r.bd}>{bd}</td>
                      </tr>
                    )
                  })
                )}
              </tbody>
            </table>
          </div>

          <div className="pager">
            <button
              type="button"
              className="btn secondary sm"
              disabled={page <= 0}
              onClick={() => setPage((p) => Math.max(0, p - 1))}
            >
              ← Prev
            </button>
            <span>
              Page {page + 1} / {totalPages}
            </span>
            <button
              type="button"
              className="btn secondary sm"
              disabled={page >= totalPages - 1}
              onClick={() => setPage((p) => p + 1)}
            >
              Next →
            </button>
          </div>
        </div>

        <DetailPanel record={selected} onClose={() => setSelectedId(null)} />
      </main>
    </div>
  )
}