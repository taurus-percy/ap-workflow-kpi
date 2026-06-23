# A/P Workflow KPI Generator

A React-based tool for generating formatted A/P workflow performance reports from Yardi Voyager exports.

## Features

- Drop-and-go Yardi export processing
- Instant KPI calculations (PM Approval, Final Approval turnaround)
- Formatted Excel output with Dashboard + Raw Data sheets
- Client-side processing — no data uploaded anywhere
- Branded for Taurus Commercial Real Estate Services

## Setup & Deployment

### Local Development

1. Clone the repo
2. `npm install`
3. `npm run dev` (runs on http://localhost:5173)

### Deploy to Vercel

1. Push to GitHub
2. Import repo in Vercel dashboard
3. Vercel auto-builds and deploys
4. Share the live URL with your team

## File Structure

```
workflow-kpi-tool/
├── index.html
├── package.json
├── vite.config.js
├── vercel.json
├── src/
│   ├── main.jsx
│   └── App.jsx
└── .gitignore
```

## Usage

1. Export your weekly A/P workflow data from Yardi Voyager
2. Drop the .xlsx file on the web page
3. Report downloads automatically

## Technologies

- React 18
- Vite (build tool)
- SheetJS (Excel processing)
