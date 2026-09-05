# chart-ooxml-adapter documentation

This library converts chart library configs into a DemoMacro-compatible Office Open JSON chart object. It does not generate Office packages.

## Guides

- [Architecture](architecture/overview.md) — pipeline, boundaries, and why the output is JSON rather than XML or PPTX
- [User guide: converting charts](user-guide/converting-charts.md) — how to call the public API
- [Developer guide: UniversalChartModel](developer-guide/universal-chart-model.md)
- [Developer guide: Office Open JSON](developer-guide/office-open-json.md)
- [Developer guide: testing](developer-guide/testing.md)

## Public contract

| Input | Output |
| --- | --- |
| Chart.js / ECharts / Highcharts / amCharts config, or `UniversalChartModel` | `OfficeOpenChart` JSON object (`type: "chart"`) |

Canonical JSON Schema: [`schemas/office-open-chart.schema.json`](../schemas/office-open-chart.schema.json)

This library does not create ZIP packages, relationships, media, `document.xml`, slide layouts, or fallback images.
