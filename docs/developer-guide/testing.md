# Testing

Run the suite from the repository root:

```bash
npm test
npm run typecheck
npm run build
```

The tests cover:

1. `UniversalChartModel` → Office Open JSON, including the documented bar example
2. Office Open JSON against `schemas/office-open-chart.schema.json` (Ajv draft 2020-12)
3. Presence of required OOXML ChartML entities (`chartSpace`, `plotArea`, chart-type element, `ser`, axes where required)
4. Source adapters for Chart.js, ECharts, Highcharts, and amCharts
5. Rejection of XML-string payloads and incomplete objects

Unit tests live under `tests/`. They do not create Office packages.
