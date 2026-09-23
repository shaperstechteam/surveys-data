# Property Records Viewer (React + Vite)

Search and browse ~7,400 commercial property records in one UI.

## Features

- Full-text search (owner, city, address, amount, brands, phone, notes, ...)
- Filters: City, Rent / Sale, Has price
- Sortable list columns
- Click any row to open detail panel with **all fields and labels**
- Export filtered results as CSV
- Data loaded from `public/properties.json`

## Setup

```bash
cd properties-viewer
npm install
npm run dev
```

Open the URL Vite prints (usually http://localhost:5173).

## Build for production

```bash
npm run build
npm run preview
```

## Project layout

```
properties-viewer/
  public/
    properties.json   # all records
  src/
    App.jsx           # main UI
    App.css
    index.css
    main.jsx
  package.json
  vite.config.js
  index.html
```
