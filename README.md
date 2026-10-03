# SMS Route Control Center

A frontend-only React/Vite prototype for an internal SMS testing operations panel.

## What is included

- Dashboard with operational KPIs and charts
- Shift Testing view with 3 shifts
- Grouping of multiple individual attempts into one operational row
- 2+ delivered attempts = Working evaluation
- Results Explorer with raw-attempt grouping
- Sales Management for recurring testing requirements
- Destination configuration
- Activity/audit log
- Excel / XLS / CSV import using SheetJS
- Local persistence using browser localStorage
- Responsive NOC-style UI
- No backend or external database required

## Grouping rule used by the prototype

An imported attempt is grouped into one shift record using:

`Country + Network + MCC + MNC + Supplier + Route Type + Sender ID`

The supplied testing export contains MCC/MNC, Supplier, Route Type, Sender ID, DLR status and delays. It does not contain a separate `SID` column, so the prototype does not invent one.

## Run

```bash
npm install
npm run dev
```

Then open the local Vite URL shown in the terminal.

## Important

This is intentionally a demo/prototype. It stores data in the browser only. It is suitable for a CEO/development-team presentation, not for production use or sensitive operational data.
