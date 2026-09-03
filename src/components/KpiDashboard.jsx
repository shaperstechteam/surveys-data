import { formatMoney, formatNumber, formatPct } from '../assets/kpis'
import './KpiDashboard.css'

function StatCard({ label, value, sub }) {
  return (
    <div className="kpi-card">
      <div className="kpi-card-label">{label}</div>
      <div className="kpi-card-value">{value}</div>
      {sub ? <div className="kpi-card-sub">{sub}</div> : null}
    </div>
  )
}

function RankTable({ title, rows, nameKey = 'name', valueKey = 'count', showPct = true }) {
  if (!rows?.length) {
    return (
      <div className="kpi-panel">
        <h3>{title}</h3>
        <p className="kpi-empty">No data</p>
      </div>
    )
  }
  const max = Math.max(...rows.map((r) => r[valueKey] || 0), 1)
  return (
    <div className="kpi-panel">
      <h3>{title}</h3>
      <div className="kpi-rank-list">
        {rows.map((r) => (
          <div className="kpi-rank-row" key={String(r[nameKey])}>
            <div className="kpi-rank-name" title={r[nameKey]}>
              {r[nameKey]}
            </div>
            <div className="kpi-rank-bar-wrap">
              <div
                className="kpi-rank-bar"
                style={{ width: `${((r[valueKey] || 0) / max) * 100}%` }}
              />
            </div>
            <div className="kpi-rank-val">
              {formatNumber(r[valueKey])}
              {showPct && r.pct != null ? (
                <span className="kpi-rank-pct"> {formatPct(r.pct)}</span>
              ) : null}
            </div>
          </div>
        ))}
      </div>
    </div>
  )
}

function CompletenessTable({ rows }) {
  return (
    <div className="kpi-panel">
      <h3>Data completeness (% filled)</h3>
      <div className="kpi-rank-list">
        {rows.map((r) => (
          <div className="kpi-rank-row" key={r.label}>
            <div className="kpi-rank-name">{r.label}</div>
            <div className="kpi-rank-bar-wrap">
              <div
                className="kpi-rank-bar kpi-bar-green"
                style={{ width: `${Math.min(r.pct, 100)}%` }}
              />
            </div>
            <div className="kpi-rank-val">
              {formatPct(r.pct)}
              <span className="kpi-rank-pct"> ({formatNumber(r.count)})</span>
            </div>
          </div>
        ))}
      </div>
    </div>
  )
}

/**
 * @param {{ kpis: object, filtered: boolean }} props
 */
export default function KpiDashboard({ kpis, filtered }) {
  if (!kpis) return null
  const a = kpis.amount

  return (
    <section className="kpi-dashboard">
      <div className="kpi-header">
        <h2>
          KPI Dashboard
          {filtered ? <span className="kpi-badge">Filtered view</span> : null}
        </h2>
        <p className="kpi-hint">
          Metrics update with your search and filters. Based on{' '}
          <strong>{formatNumber(kpis.total)}</strong> records.
        </p>
      </div>

      {/* 1. Volume overview */}
      <div className="kpi-section-title">1. Volume / overview</div>
      <div className="kpi-cards">
        <StatCard label="Total records" value={formatNumber(kpis.total)} />
        <StatCard
          label="With Amount"
          value={formatNumber(kpis.withAmount)}
          sub={formatPct(kpis.total ? (kpis.withAmount / kpis.total) * 100 : 0)}
        />
        <StatCard
          label="With Pictures"
          value={formatNumber(kpis.withPic)}
          sub={formatPct(kpis.total ? (kpis.withPic / kpis.total) * 100 : 0)}
        />
        <StatCard
          label="With Videos"
          value={formatNumber(kpis.withVid)}
          sub={formatPct(kpis.total ? (kpis.withVid / kpis.total) * 100 : 0)}
        />
        <StatCard
          label="With Coordinates"
          value={formatNumber(kpis.withCoords)}
          sub={formatPct(kpis.total ? (kpis.withCoords / kpis.total) * 100 : 0)}
        />
      </div>
      <div className="kpi-grid">
        <RankTable title="By City" rows={kpis.byCity} />
        <RankTable title="Available For (Rent / Sale)" rows={kpis.byAvailFor} />
        <RankTable title="Available (Not Rented)" rows={kpis.byNotRented} />
        <RankTable title="Property Type" rows={kpis.byPropType} />
      </div>

      {/* 2. Owner / contact */}
      <div className="kpi-section-title">2. Owner / contact</div>
      <div className="kpi-cards">
        <StatCard label="Unique owners" value={formatNumber(kpis.uniqueOwners)} />
        <StatCard
          label="Missing owner name"
          value={formatNumber(kpis.missingOwner)}
          sub={formatPct(kpis.total ? (kpis.missingOwner / kpis.total) * 100 : 0)}
        />
        <StatCard
          label="Missing contact number"
          value={formatNumber(kpis.missingPhone)}
          sub={formatPct(kpis.total ? (kpis.missingPhone / kpis.total) * 100 : 0)}
        />
      </div>
      <div className="kpi-grid">
        <RankTable title="Top owners (by listings)" rows={kpis.topOwners} showPct={false} />
      </div>

      {/* 3. Location */}
      <div className="kpi-section-title">3. Location</div>
      <div className="kpi-grid">
        <RankTable title="Top cities" rows={kpis.byCity} />
        <RankTable title="Top areas — Lahore" rows={kpis.areasLahore} />
        <RankTable title="Top areas — Islamabad" rows={kpis.areasIslamabad} />
        <RankTable title="Top areas — Rawalpindi" rows={kpis.areasRawalpindi} />
        <RankTable title="Top areas — Karachi" rows={kpis.areasKarachi} />
      </div>

      {/* 4. Deal / pricing */}
      <div className="kpi-section-title">4. Deal / pricing</div>
      <div className="kpi-cards">
        <StatCard label="Avg Amount" value={formatMoney(a.avg)} sub={`${formatNumber(a.count)} priced`} />
        <StatCard label="Median Amount" value={formatMoney(a.median)} />
        <StatCard label="Min Amount" value={formatMoney(a.min)} />
        <StatCard label="Max Amount" value={formatMoney(a.max)} />
        <StatCard
          label="With Advance"
          value={formatNumber(kpis.withAdvance)}
          sub={formatPct(kpis.total ? (kpis.withAdvance / kpis.total) * 100 : 0)}
        />
        <StatCard
          label="With Security"
          value={formatNumber(kpis.withSecurity)}
          sub={formatPct(kpis.total ? (kpis.withSecurity / kpis.total) * 100 : 0)}
        />
      </div>
      <div className="kpi-grid">
        <RankTable title="Amount distribution" rows={a.bands} />
        <RankTable title="Deal type mix" rows={kpis.byDeal} />
      </div>
      {a.byCity?.length ? (
        <div className="kpi-panel kpi-panel-wide">
          <h3>Amount by city (top cities with prices)</h3>
          <div className="kpi-table-wrap">
            <table className="kpi-table">
              <thead>
                <tr>
                  <th>City</th>
                  <th>Priced</th>
                  <th>Avg</th>
                  <th>Median</th>
                  <th>Min</th>
                  <th>Max</th>
                </tr>
              </thead>
              <tbody>
                {a.byCity.map((r) => (
                  <tr key={r.city}>
                    <td>{r.city}</td>
                    <td>{formatNumber(r.count)}</td>
                    <td>{formatMoney(r.avg)}</td>
                    <td>{formatMoney(r.median)}</td>
                    <td>{formatMoney(r.min)}</td>
                    <td>{formatMoney(r.max)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ) : null}

      {/* 5. Business Developer */}
      <div className="kpi-section-title">5. Business Developer</div>
      <div className="kpi-grid">
        <RankTable title="Records per BD" rows={kpis.byBd} />
      </div>
      {kpis.bdCityTop?.length ? (
        <div className="kpi-panel kpi-panel-wide">
          <h3>BD × City (top pairs)</h3>
          <div className="kpi-table-wrap">
            <table className="kpi-table">
              <thead>
                <tr>
                  <th>Business Developer</th>
                  <th>City</th>
                  <th>Records</th>
                </tr>
              </thead>
              <tbody>
                {kpis.bdCityTop.map((r) => (
                  <tr key={`${r.bd}-${r.city}`}>
                    <td>{r.bd}</td>
                    <td>{r.city}</td>
                    <td>{formatNumber(r.count)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ) : null}

      {/* 6. Brands */}
      <div className="kpi-section-title">6. Brands / competition</div>
      <div className="kpi-grid">
        <RankTable
          title="Top Select Brands"
          rows={kpis.brands.list}
          showPct={false}
        />
      </div>

      {/* 7. Data quality */}
      <div className="kpi-section-title">7. Data quality</div>
      <div className="kpi-cards">
        <StatCard
          label="Missing City"
          value={formatNumber(kpis.missingCity)}
          sub={formatPct(kpis.total ? (kpis.missingCity / kpis.total) * 100 : 0)}
        />
        <StatCard
          label="Missing Owner"
          value={formatNumber(kpis.missingOwner)}
          sub={formatPct(kpis.total ? (kpis.missingOwner / kpis.total) * 100 : 0)}
        />
        <StatCard
          label="Missing Amount"
          value={formatNumber(kpis.missingAmount)}
          sub={formatPct(kpis.total ? (kpis.missingAmount / kpis.total) * 100 : 0)}
        />
      </div>
      <div className="kpi-grid">
        <CompletenessTable rows={kpis.completeness} />
        <RankTable title="Entries by month" rows={kpis.byMonth} showPct={false} />
      </div>
    </section>
  )
}
