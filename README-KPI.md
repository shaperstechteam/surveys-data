# KPI files — how to add to your React viewer

## Files to copy

```
src/
  utils/
    kpis.js              ← KPI calculations
  components/
    KpiDashboard.jsx     ← KPI UI
    KpiDashboard.css     ← styles
```

## 1. Copy the three files into your project

- `src/utils/kpis.js`
- `src/components/KpiDashboard.jsx`
- `src/components/KpiDashboard.css`

## 2. Wire into `App.jsx`

Add imports:

```jsx
import { useMemo } from 'react' // if not already
import { computeKpis } from './utils/kpis'
import KpiDashboard from './components/KpiDashboard'
```

After you compute `filtered` (your filtered records array), add:

```jsx
const [showKpis, setShowKpis] = useState(true)

const kpis = useMemo(() => computeKpis(filtered), [filtered])
const isFiltered = filtered.length !== all.length
```

In the toolbar, add a toggle button:

```jsx
<button
  type="button"
  className="btn secondary"
  onClick={() => setShowKpis((v) => !v)}
>
  {showKpis ? 'Hide KPIs' : 'Show KPIs'}
</button>
```

Render the dashboard **above** the list (e.g. right after the toolbar):

```jsx
{showKpis ? <KpiDashboard kpis={kpis} filtered={isFiltered} /> : null}
```

## 3. Record keys used

KPIs expect compact JSON keys from `properties.json`:

| Key | Meaning |
|-----|---------|
| `ci` | City |
| `on` | Owner name |
| `cn` | Contact number |
| `bd` | Business developer |
| `af` | Available for (Rent/Sale) |
| `an` | Available not rented |
| `pt` | Property type |
| `am` | Amount |
| `dl` | Deal |
| `av` | Advance |
| `sc` | Security |
| `al` `ai` `ar` `ak` | Area Lahore / Islamabad / Rawalpindi / Karachi |
| `sb` `eb` `ob` | Select / Earlier / Other brands |
| `pic` `vid` | Picture / Video URL |
| `lng` `lat` | Coordinates |
| `dc` | Date created |
| `ta` `ad` | Total area / Address |

If your JSON uses full column names instead, change the keys inside `kpis.js`.

## 4. Behaviour

- KPIs recalculate on **filtered** data (search + city + rent/sale + price filters).
- Badge **Filtered view** appears when the list is not the full dataset.
