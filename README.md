# chart-ooxml-adapter

Convert JavaScript chart library configurations into a **DemoMacro-compatible Office Open JSON** chart object.

This library does **not** create PPTX, XLSX, or DOCX files. It stops at the Office Open JSON layer.

```text
ECharts / amCharts / Highcharts / Chart.js
        │
        ▼
  Chart Adapter
        │
        ▼
  UniversalChartModel
        │
        ▼
  Office Open JSON (DemoMacro)
        │
        ▼
  Office Package Generator (other libraries)
        │
        ▼
  PPTX / XLSX / DOCX
```

## Install

```bash
npm install chart-ooxml-adapter
```

## Quick start

```ts
import { convertChart } from "chart-ooxml-adapter";

const chart = convertChart({
  type: "bar",
  categories: ["A", "B", "C"],
  series: [{ name: "Sales", values: [10, 20, 30] }],
});

// chart.type === "chart"
// chart.chartSpace.chart.plotArea.barChart.ser[0].tx === "Sales"
```

Pass Chart.js, ECharts, Highcharts, or amCharts options to the same function. Detection is automatic.

Full documentation: [docs/index.md](docs/index.md)
