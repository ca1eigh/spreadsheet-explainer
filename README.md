# SheetScan - Spreadsheet Explainer

SheetScan is a React web app for finding, explaining, and fixing spreadsheet data issues.  
You can upload CSV/Excel files, inspect anomalies, edit cells inline, track version history, and export both updated data and stakeholder-ready audit reports.

## Features

- Upload `.csv`, `.xlsx`, or `.xls` files
- Automatic anomaly detection (critical, warning, missing, invalid)
- Interactive explanation panel with:
  - issue reasoning
  - same-row and same-column context
  - dependency graph visualization
  - suggested remediation guidance
- Inline cell editing (double-click any cell)
- Version history snapshots with restore support
- Download edited dataset as:
  - Excel (`.xlsx`)
  - CSV (`.csv`)
- Download shareable audit reports as:
  - PDF (`.pdf`)
  - DOC (`.doc`)

## Run Locally

### 1) Install dependencies

```bash
npm install
```

### 2) Start development server

```bash
npm run dev
```

Then open the local URL shown in your terminal (usually `http://localhost:5173`).

## Build for Production

```bash
npm run build
```

## Preview Production Build

```bash
npm run preview
```
